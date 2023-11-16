/* eslint-disable no-unused-vars */
// eslint-disable-next-line no-unused-vars
import React, { useRef } from "react";
import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const tableStyle = {
  width: "100%",
  borderCollapse: "collapse",
  margin: "20px 0",
};

const cellStyle = {
  border: "1px solid #ddd",
  padding: "8px",
};

const headerCellStyle = {
  ...cellStyle,
  backgroundColor: "#00f",
  fontWeight: "bold",
};

const inputStyle = {
  border: "none",
  padding: "8px",
  width: "100%",
  boxSizing: "border-box",
};

const buttonStyle = {
  padding: "10px 20px",
  fontSize: "16px",
  marginTop: "20px",
  cursor: "pointer",
  backgroundColor: "#4CAF50",
  color: "white",
  border: "none",
  borderRadius: "4px",
};

function ExcelImportExport() {
  const [data, setData] = useState([]);
  const [editingData, setEditingData] = useState([]);
  const [editedFileName, setEditedFileName] = useState("");
  const fileInputRef = useRef(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const binaryData = e.target.result;
        const workbook = XLSX.read(binaryData, { type: "binary" });

        const sheetName = workbook.SheetNames[0];
        const excelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        setData(excelData);
        setEditingData([...excelData]);
      };
      reader.readAsBinaryString(file);
    }
  };

  const handleEdit = (rowIndex, columnName, value) => {
    setEditingData((prevEditingData) => {
      const updatedData = [...prevEditingData];

      console.log(updatedData, "updatedData");
      console.log(rowIndex, "rowIndex");
      console.log(columnName, "columnName");
      console.log(value, "value");

      updatedData[rowIndex][columnName] = value;
      return updatedData;
    });
  };

  const handleSave = async () => {
    const editedWorkbook = new ExcelJS.Workbook();
    const editedSheet = editedWorkbook.addWorksheet("Sheet1");

    // Table headers
    editedSheet.columns = [
      { header: "№ п/п", key: "Number", style: headerCellStyle },
      { header: "Наименование", key: "Name", style: headerCellStyle },
      { header: "Вид запасов", key: "Type", style: headerCellStyle },
      { header: "Ед. изм", key: "Unit", style: headerCellStyle },
      { header: "Количество", key: "Quantity", style: headerCellStyle },
      { header: "Цена", key: "Price", style: headerCellStyle },
      { header: "Сумма", key: "Amount", style: headerCellStyle },
      {
        header: "Дата приобретения",
        key: "PurchaseDate",
        style: headerCellStyle,
      },
      { header: "Поставщик", key: "Supplier", style: headerCellStyle },
      { header: "ИНН", key: "INN", style: headerCellStyle },
      { header: "Примечание", key: "Note", style: headerCellStyle },
      { header: "Предложение", key: "Offer", style: headerCellStyle },
    ];

    editingData.forEach((row) => {
      editedSheet.addRow(row);
    });

    const editedFileData = await editedWorkbook.xlsx.writeBuffer();

    const blob = new Blob([editedFileData], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = editedFileName || "edited_data.xlsx";
    a.click();
  };

  return (
    <div>
      <h1>Edit Excel Data</h1>
      <input
        style={buttonStyle}
        type="file"
        accept=".xls, .xlsx"
        onChange={handleFileChange}
      />
      <br />
      <button style={buttonStyle} onClick={handleSave}>
        Save to Excel
      </button>
      {data.length > 0 && (
        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={headerCellStyle}>№ п/п</th>
              <th style={headerCellStyle}>Наименование</th>
              <th style={headerCellStyle}>Вид запасов</th>
              <th style={headerCellStyle}>Ед. изм</th>
              <th style={headerCellStyle}>Количество</th>
              <th style={headerCellStyle}>Цена</th>
              <th style={headerCellStyle}>Сумма</th>
              <th style={headerCellStyle}>Дата приобретения</th>
              <th style={headerCellStyle}>Поставщик</th>
              <th style={headerCellStyle}>ИНН</th>
              <th style={headerCellStyle}>Примечание</th>
              <th style={headerCellStyle}>Предложение</th>
            </tr>
          </thead>
          <tbody>
            {editingData.map((row, rowIndex) => (
              <tr key={rowIndex}>
                <td style={cellStyle}>{row["№ п/п"]}</td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Наименование"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Наименование", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Вид запасов"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Вид запасов", e.target.value)
                    }
                  />
                </td>

                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Ед. изм"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Ед. изм", e.target.value)
                    }
                  />
                </td>

                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Количество"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Количество", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Цена"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Цена", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Сумма"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Сумма", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Дата приобретения"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Дата приобретения", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Поставщик"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Поставщик", e.target.value)
                    }
                  />
                </td>

                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["ИНН"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "ИНН", e.target.value)
                    }
                  />
                </td>

                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Примечание"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Примечание", e.target.value)
                    }
                  />
                </td>
                <td style={cellStyle}>
                  <input
                    style={inputStyle}
                    type="text"
                    value={row["Предложение"]}
                    onChange={(e) =>
                      handleEdit(rowIndex, "Предложение", e.target.value)
                    }
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default ExcelImportExport;
