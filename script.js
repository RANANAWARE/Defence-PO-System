"use strict";

let extractedData = [];

const fileInput = document.getElementById("pdfFile");
const extractButton = document.getElementById("extractButton");
const downloadButton = document.getElementById("downloadButton");
const clearButton = document.getElementById("clearButton");
const summaryText = document.getElementById("summaryText");
const tableBody = document.querySelector("#resultTable tbody");

if (typeof pdfjsLib !== "undefined") {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

extractButton.addEventListener("click", extractData);
downloadButton.addEventListener("click", downloadExcel);
clearButton.addEventListener("click", clearResults);
fileInput.addEventListener("change", () => {
    const count = fileInput.files.length;
    setSummary(count ? `${count} PDF file${count === 1 ? "" : "s"} selected.` : "No PDF selected yet.");
});

async function extractData() {
    const files = Array.from(fileInput.files || []);

    if (!files.length) {
        alert("Please select at least one PDF file first.");
        return;
    }

    if (typeof pdfjsLib === "undefined") {
        setSummary("PDF.js did not load. Check your internet connection or library link.", true);
        return;
    }

    extractedData = [];
    displayTable([]);
    setBusy(true);

    const failures = [];

    try {
        for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
            const file = files[fileIndex];
            setSummary(`Reading ${fileIndex + 1} of ${files.length}: ${file.name}`);

            try {
                const text = await readPdfText(file);
                const poNumber = extractPONumber(text);
                const rows = extractItems(text, poNumber, file.name);

                if (!rows.length) {
                    failures.push(`${file.name}: no line items found`);
                } else {
                    extractedData.push(...rows);
                }
            } catch (error) {
                console.error(`Failed to process ${file.name}`, error);
                failures.push(`${file.name}: ${error.message || "unable to read PDF"}`);
            }
        }

        extractedData.sort((a, b) => {
            const poCompare = String(a["PO Number"]).localeCompare(String(b["PO Number"]));
            return poCompare || Number(a["PO Line"]) - Number(b["PO Line"]);
        });

        displayTable(extractedData);
        downloadButton.disabled = extractedData.length === 0;

        const successMessage = `${extractedData.length} row${extractedData.length === 1 ? "" : "s"} extracted from ${files.length} PDF file${files.length === 1 ? "" : "s"}.`;
        setSummary(failures.length ? `${successMessage} ${failures.join(" | ")}` : successMessage, failures.length > 0);

        if (!extractedData.length) {
            alert("No line items were found. The PDF may be scanned as an image or use an unsupported layout.");
        }
    } finally {
        setBusy(false);
    }
}

async function readPdfText(file) {
    const data = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data }).promise;
    const pages = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        pages.push(buildPageText(content.items));
    }

    return normaliseText(pages.join("\n"));
}

// Reconstruct each page in visual reading order. This is more reliable than
// simply joining PDF.js items because PDFs do not always store text in row order.
function buildPageText(items) {
    const tokens = items
        .filter(item => item && typeof item.str === "string" && item.str.trim())
        .map(item => ({
            text: item.str.trim(),
            x: Number(item.transform?.[4] || 0),
            y: Number(item.transform?.[5] || 0),
            width: Number(item.width || 0)
        }))
        .sort((a, b) => {
            const sameLine = Math.abs(a.y - b.y) <= 2.5;
            return sameLine ? a.x - b.x : b.y - a.y;
        });

    const lines = [];
    for (const token of tokens) {
        let line = lines.find(candidate => Math.abs(candidate.y - token.y) <= 2.5);
        if (!line) {
            line = { y: token.y, tokens: [] };
            lines.push(line);
        }
        line.tokens.push(token);
    }

    return lines
        .sort((a, b) => b.y - a.y)
        .map(line => line.tokens.sort((a, b) => a.x - b.x).map(token => token.text).join(" "))
        .join("\n");
}

function normaliseText(value) {
    return String(value || "")
        .replace(/[\u00A0\u2007\u202F]/g, " ")
        .replace(/[–—]/g, "-")
        .replace(/\r/g, "\n")
        .replace(/[ \t]+/g, " ")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

function extractPONumber(text) {
    const labelled = text.match(/\bOrder\s*No\.?\s*:?\s*(\d{7,12})\b/i);
    if (labelled) return labelled[1];

    const fallback = text.match(/\b3000\d{6,}\b/);
    return fallback ? fallback[0] : "";
}

function extractItems(text, poNumber, fileName) {
    const rows = [];
    const flatText = normaliseText(text).replace(/\n/g, " ");

    // Item numbers are immediately followed by "Material:" in these Defence POs.
    // The lookahead ends a block at the next line item or at the order totals.
    const blockPattern = /(?:^|\s)(\d{1,6})\s+Material\s*:\s*([A-Z0-9][A-Z0-9 .\/-]*?)(?=\s+\S)[\s\S]*?(?=(?:\s+\d{1,6}\s+Material\s*:)|(?:\s+Total\s+Order\s+Net\s+Value\s*:)|$)/gi;
    let blockMatch;

    while ((blockMatch = blockPattern.exec(flatText)) !== null) {
        const itemNo = blockMatch[1];
        const block = blockMatch[0].trim();
        const header = parseItemHeader(block);

        if (!header) {
            console.warn("Could not parse item header:", block);
            continue;
        }

        const nsn = extractNSN(block) || header.materialNumber;
        const partNumber = extractMercedesPartNumber(block);

        rows.push({
            "PO Number": poNumber,
            "PO Line": itemNo,
            "Part Number": partNumber,
            "NSN": nsn,
            "Description": header.description,
            "Qty": header.quantity,
            "CoA $": header.itemTotal,
            "EDD": header.deliveryDate,
            "Source PDF": fileName
        });
    }

    return rows;
}

function parseItemHeader(block) {
    // Capture from "Material:" through the AUD row. Values after EA are:
    // Unit Price, GST and Item Total (incl. GST).
    const pattern = /Material\s*:\s*([A-Z0-9][A-Z0-9 .\/-]*?)\s+(.+?)\s+(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4})\s+([\d,]+(?:\.\d+)?)\s+(?:EA|EACH)\s+([\d,]+(?:\.\d{1,4})?)\s+([\d,]+(?:\.\d{1,2})?)\s+([\d,]+(?:\.\d{1,2})?)\s+AUD\b/i;
    const match = block.match(pattern);

    if (!match) return null;

    return {
        materialNumber: cleanIdentifier(match[1]),
        description: cleanText(match[2]),
        deliveryDate: normaliseDate(match[3]),
        quantity: match[4].replace(/,/g, ""),
        unitPrice: match[5],
        gst: match[6],
        itemTotal: match[7]
    };
}

function extractNSN(block) {
    const match = block.match(/\bNSN\s*:\s*([A-Z0-9][A-Z0-9 .\/-]*?)(?=\s+Manufacturer\s+Part\s+No\.?\s*:|\s+\d{1,6}\s+Material\s*:|\s+Total\s+Order|$)/i);
    return match ? cleanIdentifier(match[1]) : "";
}

function extractMercedesPartNumber(block) {
    const labelMatch = block.match(/Manufacturer\s+Part\s+No\.?\s*:\s*([\s\S]*?)(?=\s+\d{1,6}\s+Material\s*:|\s+Total\s+Order\s+Net|$)/i);
    if (!labelMatch) return "";

    const entries = labelMatch[1]
        .split(/\s*,\s*/)
        .map(entry => entry.trim())
        .filter(Boolean);

    // Prefer the number explicitly associated with Mercedes-Benz.
    const mercedesEntry = entries.find(entry => /MERCEDES[- ]BENZ/i.test(entry));
    const chosenEntry = mercedesEntry || entries[0] || "";
    const numberPart = chosenEntry.split(/\s+\/\s+/)[0];

    return cleanPartNumber(numberPart);
}

function cleanIdentifier(value) {
    return cleanText(value).replace(/\s+/g, "");
}

function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
}

function cleanPartNumber(value) {
    return cleanText(value).replace(/\s+/g, "").toUpperCase();
}

function normaliseDate(value) {
    const parts = String(value).split(/[\/.-]/);
    return parts.length === 3 ? `${parts[0].padStart(2, "0")}/${parts[1].padStart(2, "0")}/${parts[2]}` : value;
}

function displayTable(data) {
    tableBody.replaceChildren();

    for (const row of data) {
        const tr = document.createElement("tr");
        const columns = ["PO Number", "PO Line", "Part Number", "NSN", "Description", "Qty", "CoA $", "EDD", "Source PDF"];

        for (const column of columns) {
            const td = document.createElement("td");
            td.textContent = row[column] ?? "";
            tr.appendChild(td);
        }
        tableBody.appendChild(tr);
    }
}

function downloadExcel() {
    if (!extractedData.length) {
        alert("Please extract data first.");
        return;
    }

    if (typeof XLSX === "undefined") {
        alert("The Excel library did not load. Check your internet connection or the SheetJS link.");
        return;
    }

    const columns = ["PO Number", "PO Line", "Part Number", "NSN", "Description", "Qty", "CoA $", "EDD", "Source PDF"];
    const excelRows = extractedData.map(row => {
        const result = {};
        for (const column of columns) result[column] = row[column] ?? "";
        result["Qty"] = toNumber(row["Qty"]);
        result["CoA $"] = toNumber(row["CoA $"]);
        return result;
    });

    const worksheet = XLSX.utils.json_to_sheet(excelRows, { header: columns });
    worksheet["!cols"] = [
        { wch: 15 }, { wch: 10 }, { wch: 22 }, { wch: 16 }, { wch: 38 },
        { wch: 10 }, { wch: 14 }, { wch: 13 }, { wch: 28 }
    ];
    worksheet["!autofilter"] = { ref: worksheet["!ref"] };
    worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };

    for (let row = 2; row <= excelRows.length + 1; row++) {
        if (worksheet[`F${row}`]) worksheet[`F${row}`].z = "0";
        if (worksheet[`G${row}`]) worksheet[`G${row}`].z = "$#,##0.00";
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "PO Data");

    const uniquePOs = [...new Set(extractedData.map(row => row["PO Number"]).filter(Boolean))];
    const outputName = uniquePOs.length === 1 ? `${uniquePOs[0]}.xlsx` : "PO_Data_Combined.xlsx";
    XLSX.writeFile(workbook, outputName);
}

function toNumber(value) {
    const numeric = Number(String(value ?? "").replace(/[$,\s]/g, ""));
    return Number.isFinite(numeric) ? numeric : value;
}

function setBusy(isBusy) {
    extractButton.disabled = isBusy;
    fileInput.disabled = isBusy;
    clearButton.disabled = isBusy;
    extractButton.textContent = isBusy ? "Extracting..." : "Extract Data";
}

function setSummary(message, isError = false) {
    summaryText.textContent = message;
    summaryText.classList.toggle("error", isError);
}

function clearResults() {
    fileInput.value = "";
    extractedData = [];
    displayTable([]);
    downloadButton.disabled = true;
    setSummary("No PDF selected yet.");
}
