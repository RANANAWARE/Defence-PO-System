let extractedData = [];

async function extractData() {

    const file = document.getElementById("pdfFile").files[0];

    if (!file) {
        alert("Please select a PDF file first.");
        return;
    }

    try {

        const arrayBuffer = await file.arrayBuffer();

        const pdf = await pdfjsLib.getDocument({
            data: arrayBuffer
        }).promise;

        let fullText = "";

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {

            const page = await pdf.getPage(pageNumber);
            const textContent = await page.getTextContent();

            const pageText = textContent.items
                .map(item => item.str)
                .join(" ");

            fullText += " " + pageText;
        }

        console.log("PDF TEXT:");
        console.log(fullText);

        extractedData = [];

        const poNumber = extractPONumber(fullText);
        const items = extractItems(fullText, poNumber, file.name);

        extractedData = items;

        displayTable(extractedData);

        document.getElementById("summaryText").innerText =
            `${extractedData.length} rows extracted from ${file.name}`;

        if (extractedData.length === 0) {
            alert("No records found. Check if the PDF text format is readable.");
        } else {
            alert(`${extractedData.length} records extracted successfully.`);
        }

    } catch (error) {

        console.error(error);
        alert("Unable to read PDF. Please open F12 Console and check the error.");
    }
}

function extractPONumber(text) {

    const match = text.match(/Order No:\s*(3000\d+)/i);

    if (match) {
        return match[1];
    }

    const fallbackMatch = text.match(/\b3000\d+\b/);

    return fallbackMatch ? fallbackMatch[0] : "";
}

function extractItems(text, poNumber, fileName) {

    const rows = [];

    const itemPattern =
        /(\d+)\s+Material:\s*(12\d+)\s+(.+?)\s+(\d{2}\/\d{2}\/\d{4})\s+([\d,]+)\s+EA\s+[\d,.]+\s+[\d,.]+\s+([\d,.]+)\s+AUD/gi;

    let itemMatches = [];
    let itemMatch;

    while ((itemMatch = itemPattern.exec(text)) !== null) {

        itemMatches.push({
            itemNo: itemMatch[1],
            nsnFromMaterial: itemMatch[2],
            description: cleanText(itemMatch[3]),
            deliveryDate: itemMatch[4],
            quantity: itemMatch[5].replace(/,/g, ""),
            coaValue: itemMatch[6],
            index: itemMatch.index
        });

    }

    const manufacturerPattern =
        /NSN:\s*(12\d+)\s*Manufacturer Part No:\s*([^\/,]+)\s*\//gi;

    let manufacturerMatches = [];
    let manufacturerMatch;

    while ((manufacturerMatch = manufacturerPattern.exec(text)) !== null) {

        manufacturerMatches.push({
            nsn: manufacturerMatch[1],
            partNumber: cleanPartNumber(manufacturerMatch[2]),
            index: manufacturerMatch.index
        });

    }

    for (let i = 0; i < itemMatches.length; i++) {

        const item = itemMatches[i];

        const matchingManufacturer = manufacturerMatches.find(
            m => m.nsn === item.nsnFromMaterial
        );

        rows.push({
            "PO Number": poNumber,
            "PO Line": item.itemNo,
            "Part Number": matchingManufacturer
                ? matchingManufacturer.partNumber
                : "",
            "NSN": item.nsnFromMaterial,
            "Description": item.description,
            "Qty": item.quantity,
            "CoA $": item.coaValue,
            "EDD": item.deliveryDate,
            "Source PDF": fileName
        });
    }

    return rows;
}

function cleanText(value) {

    if (!value) {
        return "";
    }

    return value
        .replace(/\s+/g, " ")
        .trim();
}

function cleanPartNumber(value) {

    if (!value) {
        return "";
    }

    return value
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function displayTable(data) {

    const tbody = document.querySelector("#resultTable tbody");

    tbody.innerHTML = "";

    data.forEach(row => {

        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td>${row["PO Number"]}</td>
            <td>${row["PO Line"]}</td>
            <td>${row["Part Number"]}</td>
            <td>${row["NSN"]}</td>
            <td>${row["Description"]}</td>
            <td>${row["Qty"]}</td>
            <td>${row["CoA $"]}</td>
            <td>${row["EDD"]}</td>
            <td>${row["Source PDF"]}</td>
        `;

        tbody.appendChild(tr);
    });
}

function downloadExcel() {

    if (extractedData.length === 0) {
        alert("Please click Extract Data first.");
        return;
    }

    if (typeof XLSX === "undefined") {
        alert("Excel library is not loaded. Please check the SheetJS script link in index.html.");
        return;
    }

    const excelData = extractedData.map(row => ({
        "PO Number": row["PO Number"],
        "PO Line": row["PO Line"],
        "Part Number": row["Part Number"],
        "Description": row["Description"],
        "Qty": row["Qty"],
        "NSN": row["NSN"],
        "CoA $": row["CoA $"],
        "EDD": row["EDD"],
        "Source PDF": row["Source PDF"]
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);

    worksheet["!cols"] = [
        { wch: 15 },
        { wch: 10 },
        { wch: 18 },
        { wch: 35 },
        { wch: 10 },
        { wch: 15 },
        { wch: 15 },
        { wch: 25 }
    ];

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "PO Data"
    );

    let fileName = "PO_Data.xlsx";

    if (
        extractedData.length > 0 &&
        extractedData[0]["PO Number"]
    ) {
        fileName = extractedData[0]["PO Number"] + ".xlsx";
    }

    try {

        XLSX.writeFile(
            workbook,
            fileName
        );

        alert("Excel file downloaded successfully!");

    } catch (error) {

        console.error(error);

        const excelBuffer = XLSX.write(workbook, {
            bookType: "xlsx",
            type: "array"
        });

        const blob = new Blob([excelBuffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        const downloadLink = document.createElement("a");
        const url = URL.createObjectURL(blob);

        downloadLink.href = url;
        downloadLink.download = fileName;
        document.body.appendChild(downloadLink);
        downloadLink.click();

        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(url);

        alert("Excel file downloaded successfully!");
    }
}