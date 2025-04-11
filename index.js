const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// List of numbers to highlight (as strings for exact matching)
const steelTariffs = [
  "73023000",
  "73072210",
  "73072900",
  "73079150",
  "73079330",
  "73079910",
  "73081000",
  "73083050",
  "73089060",
  "73090000",
  "73102900",
  "73121010",
  "73121050",
  "73121080",
  "73130000",
  "73141230",
  "73141410",
  "73141460",
  "73142000",
  "73143900",
  "73144930",
  "73151100",
  "73152010",
  "73158210",
  "73158270",
  "73158950",
  "73170010",
  "73170055",
  "73181100",
  "73181410",
  "73181540",
  "73181580",
  "73182100",
  "73182400",
  "73194030",
  "73199090",
  "73201090",
  "73209010",
  "73211130",
  "73211900",
  "73218210",
  "73219010",
  "73219050",
  "73229000",
  "73239400",
  "73239950",
  "73241000",
  "73259100",
  "73261100",
  "73269010",
  "73269045",
  "73072110",
  "73072250",
  "73079110",
  "73079230",
  "73079360",
  "73079930",
  "73082000",
  "73084000",
  "73089070",
  "73101000",
  "73110000",
  "73121020",
  "73121060",
  "73121090",
  "73141210",
  "73141260",
  "73141420",
  "73141490",
  "73143110",
  "73144100",
  "73144960",
  "73151200",
  "73152050",
  "73158230",
  "73158910",
  "73159000",
  "73170020",
  "73170065",
  "73181200",
  "73181450",
  "73181550",
  "73181600",
  "73182200",
  "73182900",
  "73194050",
  "73201030",
  "73202010",
  "73209050",
  "73211160",
  "73218110",
  "73218250",
  "73219020",
  "73219060",
  "73231000",
  "73239910",
  "73239970",
  "73242900",
  "73259910",
  "73261900",
  "73269025",
  "73269060",
  "73072150",
  "73072300",
  "73079130",
  "73079290",
  "73079390",
  "73079950",
  "73083010",
  "73089030",
  "73089095",
  "73102100",
  "73121005",
  "73121030",
  "73121070",
  "73129000",
  "73141220",
  "73141290",
  "73141430",
  "73141901",
  "73143150",
  "73144200",
  "73145000",
  "73151900",
  "73158100",
  "73158250",
  "73158930",
  "73160000",
  "73170075",
  "73181300",
  "73181520",
  "73181560",
  "73181900",
  "73182300",
  "73194020",
  "73199010",
  "73201060",
  "73202050",
  "73211110",
  "73211200",
  "73218150",
  "73218900",
  "73219040",
  "73221900",
  "73239300",
  "73239930",
  "73239990",
  "73249000",
  "73259950",
  "73262000",
  "73269035",
  "73269086",
  "7317005501",
  "7317005511",
  "7317005520",
  "7317005550",
  "7317006530",
  "7317005502",
  "7317005518",
  "7317005530",
  "7317005570",
  "7317005508",
  "7317005519",
  "7317005540",
  "7317005590",
  "7302.30.00",
  "7307.22.10",
  "7307.29.00",
  "7307.91.50",
  "7307.93.30",
  "7307.99.10",
  "7308.10.00",
  "7308.30.50",
  "7308.90.60",
  "7309.00.00",
  "7310.29.00",
  "7312.10.10",
  "7312.10.50",
  "7312.10.80",
  "7313.00.00",
  "7314.12.30",
  "7314.14.10",
  "7314.14.60",
  "7314.20.00",
  "7314.39.00",
  "7314.49.30",
  "7315.11.00",
  "7315.20.10",
  "7315.82.10",
  "7315.82.70",
  "7315.89.50",
  "7317.00.10",
  "7317.00.55",
  "7318.11.00",
  "7318.14.10",
  "7318.15.40",
  "7318.15.80",
  "7318.21.00",
  "7318.24.00",
  "7319.40.30",
  "7319.90.90",
  "7320.10.90",
  "7320.90.10",
  "7321.11.30",
  "7321.19.00",
  "7321.82.10",
  "7321.90.10",
  "7321.90.50",
  "7322.90.00",
  "7323.94.00",
  "7323.99.50",
  "7324.10.00",
  "7325.91.00",
  "7326.11.00",
  "7326.90.10",
  "7326.90.45",
  "7307.21.10",
  "7307.22.50",
  "7307.91.10",
  "7307.92.30",
  "7307.93.60",
  "7307.99.30",
  "7308.20.00",
  "7308.40.00",
  "7308.90.70",
  "7310.10.00",
  "7311.00.00",
  "7312.10.20",
  "7312.10.60",
  "7312.10.90",
  "7314.12.10",
  "7314.12.60",
  "7314.14.20",
  "7314.14.90",
  "7314.31.10",
  "7314.41.00",
  "7314.49.60",
  "7315.12.00",
  "7315.20.50",
  "7315.82.30",
  "7315.89.10",
  "7315.90.00",
  "7317.00.20",
  "7317.00.65",
  "7318.12.00",
  "7318.14.50",
  "7318.15.50",
  "7318.16.00",
  "7318.22.00",
  "7318.29.00",
  "7319.40.50",
  "7320.10.30",
  "7320.20.10",
  "7320.90.50",
  "7321.11.60",
  "7321.81.10",
  "7321.82.50",
  "7321.90.20",
  "7321.90.60",
  "7323.10.00",
  "7323.99.10",
  "7323.99.70",
  "7324.29.00",
  "7325.99.10",
  "7326.19.00",
  "7326.90.25",
  "7326.90.60",
  "7307.21.50",
  "7307.23.00",
  "7307.91.30",
  "7307.92.90",
  "7307.93.90",
  "7307.99.50",
  "7308.30.10",
  "7308.90.30",
  "7308.90.95",
  "7310.21.00",
  "7312.10.05",
  "7312.10.30",
  "7312.10.70",
  "7312.90.00",
  "7314.12.20",
  "7314.12.90",
  "7314.14.30",
  "7314.19.01",
  "7314.31.50",
  "7314.42.00",
  "7314.50.00",
  "7315.19.00",
  "7315.81.00",
  "7315.82.50",
  "7315.89.30",
  "7316.00.00",
  "7317.00.75",
  "7318.13.00",
  "7318.15.20",
  "7318.15.60",
  "7318.19.00",
  "7318.23.00",
  "7319.40.20",
  "7319.90.10",
  "7320.10.60",
  "7320.20.50",
  "7321.11.10",
  "7321.12.00",
  "7321.81.50",
  "7321.89.00",
  "7321.90.40",
  "7322.19.00",
  "7323.93.00",
  "7323.99.30",
  "7323.99.90",
  "7324.90.00",
  "7325.99.50",
  "7326.20.00",
  "7326.90.35",
  "7326.90.86",
  "7317.00.5501",
  "7317.00.5511",
  "7317.00.5520",
  "7317.00.5550",
  "7317.00.6530",
  "7317.00.5502",
  "7317.00.5518",
  "7317.00.5530",
  "7317.00.5570",
  "7317.00.5508",
  "7317.00.5519",
  "7317.00.5540",
  "7317.00.5590",
];

const steelTariffs2 = [
  "84313100",
  "84314990",
  "85479000",
  "94059940",
  "84314200",
  "84321000",
  "94032000",
  "94062000",
  "84314910",
  "84329000",
  "94059920",
  "94069001",
  "8431.31.00",
  "8431.49.90",
  "8547.90.00",
  "9405.99.40",
  "8431.42.00",
  "8432.10.00",
  "9403.20.00",
  "9406.20.00",
  "8431.49.10",
  "8432.90.00",
  "9405.99.20",
  "9406.90.01",
];

const alumTariffs = [
  "6603908100",
  "8302106030",
  "8302200000",
  "8302413000",
  "8302416050",
  "8302423015",
  "8302496045",
  "8302500000",
  "8305100050",
  "8415908025",
  "8418998005",
  "8419505000",
  "8424909080",
  "8479899599",
  "8481909060",
  "8487900080",
  "8513902000",
  "8516908050",
  "8529907300",
  "8538100000",
  "8547900020",
  "8708103050",
  "8708806590",
  "8807300060",
  "9401999081",
  "9403991040",
  "9403999020",
  "9405994020",
  "9506516000",
  "9506910010",
  "9506990510",
  "9506991500",
  "9506992800",
  "9507302000",
  "9507308000",
  "8302103000",
  "8302106060",
  "8302303010",
  "8302416015",
  "8302416080",
  "8302423065",
  "8302496055",
  "8302603000",
  "8306300000",
  "8415908045",
  "8418998050",
  "8419901000",
  "8473302000",
  "8479908500",
  "8481909085",
  "8503009520",
  "8515902000",
  "8517710000",
  "8529909760",
  "8541900000",
  "8547900030",
  "87081060",
  "8708996890",
  "9013908000",
  "94031000",
  "9403999010",
  "9403999040",
  "9506114080",
  "9506594040",
  "9506910020",
  "9506990520",
  "9506992000",
  "9506995500",
  "9507304000",
  "9507906000",
  "8302106090",
  "8302303060",
  "8302416045",
  "8302423010",
  "8302496035",
  "8302496085",
  "8302609000",
  "8414596590",
  "8415908085",
  "8418998060",
  "8422900640",
  "8473305100",
  "8479909596",
  "8486900000",
  "8508700000",
  "8516905000",
  "8517790000",
  "8536908585",
  "8543908885",
  "8547900040",
  "8708295160",
  "8716805010",
  "9031909195",
  "94032000",
  "9403999015",
  "9403999045",
  "9506514000",
  "9506702090",
  "9506910030",
  "9506990530",
  "9506992580",
  "9506996080",
  "9507306000",
  "9603908050",
  "6603.90.8100",
  "8302.10.6030",
  "8302.20.0000",
  "8302.41.3000",
  "8302.41.6050",
  "8302.42.3015",
  "8302.49.6045",
  "8302.50.0000",
  "8305.10.0050",
  "8415.90.8025",
  "8418.99.8005",
  "8419.50.5000",
  "8424.90.9080",
  "8479.89.9599",
  "8481.90.9060",
  "8487.90.0080",
  "8513.90.2000",
  "8516.90.8050",
  "8529.90.7300",
  "8538.10.0000",
  "8547.90.0020",
  "8708.10.3050",
  "8708.80.6590",
  "8807.30.0060",
  "9401.99.9081",
  "9403.99.1040",
  "9403.99.9020",
  "9405.99.4020",
  "9506.51.6000",
  "9506.91.0010",
  "9506.99.0510",
  "9506.99.1500",
  "9506.99.2800",
  "9507.30.2000",
  "9507.30.8000",
  "8302.10.3000",
  "8302.10.6060",
  "8302.30.3010",
  "8302.41.6015",
  "8302.41.6080",
  "8302.42.3065",
  "8302.49.6055",
  "8302.60.3000",
  "8306.30.0000",
  "8415.90.8045",
  "8418.99.8050",
  "8419.90.1000",
  "8473.30.2000",
  "8479.90.8500",
  "8481.90.9085",
  "8503.00.9520",
  "8515.90.2000",
  "8517.71.0000",
  "8529.90.9760",
  "8541.90.0000",
  "8547.90.0030",
  "8708.10.60",
  "8708.99.6890",
  "9013.90.8000",
  "9403.10.00",
  "9403.99.9010",
  "9403.99.9040",
  "9506.11.4080",
  "9506.59.4040",
  "9506.91.0020",
  "9506.99.0520",
  "9506.99.2000",
  "9506.99.5500",
  "9507.30.4000",
  "9507.90.6000",
  "8302.10.6090",
  "8302.30.3060",
  "8302.41.6045",
  "8302.42.3010",
  "8302.49.6035",
  "8302.49.6085",
  "8302.60.9000",
  "8414.59.6590",
  "8415.90.8085",
  "8418.99.8060",
  "8422.90.0640",
  "8473.30.5100",
  "8479.90.9596",
  "8486.90.0000",
  "8508.70.0000",
  "8516.90.5000",
  "8517.79.0000",
  "8536.90.8585",
  "8543.90.8885",
  "8547.90.0040",
  "8708.29.5160",
  "8716.80.5010",
  "9031.90.9195",
  "9403.20.00",
  "9403.99.9015",
  "9403.99.9045",
  "9506.51.4000",
  "9506.70.2090",
  "9506.91.0030",
  "9506.99.0530",
  "9506.99.2580",
  "9506.99.6080",
  "9507.30.6000",
  "9603.90.8050",
];

const alumTariffs2 = [
  "76101000",
  "76109000",
  "7615102015",
  "7610.10.00",
  "7610.90.00",
  "7615.10.2015",
  "7615102025",
  "7615105020",
  "7615107130",
  "7615109100",
  "7616991000",
  "7616995190",
  "7615103015",
  "7615105040",
  "7615107155",
  "7615200000",
  "7616995130",
  "7615103025",
  "7615107125",
  "7615107180",
  "7616109090",
  "7616995140",
  "7615.10.2025",
  "7615.10.5020",
  "7615.10.7130",
  "7615.10.9100",
  "7616.99.1000",
  "7616.99.5190",
  "7615.10.3015",
  "7615.10.5040",
  "7615.10.7155",
  "7615.20.0000",
  "7616.99.5130",
  "7615.10.3025",
  "7615.10.7125",
  "7615.10.7180",
  "7616.10.9090",
  "7616.99.5140",
];

// Function to highlight numbers in an Excel file
async function highlightNumbers(inputFilePath, outputFilePath) {
  const workbook = new ExcelJS.Workbook();

  try {
    // Load the input Excel file
    await workbook.xlsx.readFile(inputFilePath);

    // Iterate through each worksheet
    workbook.eachSheet((worksheet) => {
      // Iterate through each row
      worksheet.eachRow((row) => {
        // Iterate through each cell in the row
        row.eachCell((cell) => {
          // Convert cell value to a string for consistent comparison
          const cellValue = cell.value?.toString() || "";

          const isMatch = steelTariffs.some((tariff) =>
            cellValue.includes(tariff)
          );
          const isMatch2 = steelTariffs2.some((tariff) =>
            cellValue.includes(tariff)
          );
          const isMatch3 = alumTariffs.some((tariff) =>
            cellValue.includes(tariff)
          );
          const isMatch4 = alumTariffs2.some((tariff) =>
            cellValue.includes(tariff)
          );
          // 99038190
          if (isMatch) {
            cell.value += " \n99038190";
            cell.font = {
              name: "Arial",
              size: 12,
              color: { argb: "00000000" }, // black text
            };
          }
          // 99038191
          if (isMatch2) {
            cell.value += " \n99038191";
            cell.font = {
              name: "Arial",
              size: 12,
              color: { argb: "00000000" }, // black text
            };
          }
          // 99038508
          if (isMatch3) {
            cell.value += " \n99038508";
            cell.font = {
              name: "Arial",
              size: 12,
              color: { argb: "00000000" }, // black text
            };
          }

          // 99038507
          if (isMatch4) {
            cell.value += " \n99038507";
            cell.font = {
              name: "Arial",
              size: 12,
              color: { argb: "00000000" }, // black text
            };
          }
        });
      });
    });

    // Save the modified workbook to a new file
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`Highlighted file saved to: ${outputFilePath}`);
  } catch (error) {
    console.error("Error processing the Excel file:", error);
  }
}

// Handle file drag-and-drop
function handleFileDrop(filePath) {
  const inputFilePath = filePath;
  const outputFilePath = path.join(
    path.dirname(inputFilePath),
    `highlighted_${path.basename(inputFilePath)}`
  );

  highlightNumbers(inputFilePath, outputFilePath);
}

// Check if a file path is provided as a command-line argument
if (process.argv[2]) {
  handleFileDrop(process.argv[2]);
} else {
  console.log("Drag and drop an Excel file onto this executable.");
}
