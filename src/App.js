import React, { useState } from "react";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";

const App = () => {
  const [formData, setFormData] = useState({});
  const [file, setFile] = useState(null);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile && selectedFile.name.endsWith(".xlsx")) {
      setFile(selectedFile);
    } else {
      alert("Please upload a valid Excel file (.xlsx)");
    }
  };

  const handleFileUpload = () => {
    if (!file) {
      alert("Please upload a file first.");
      return;
    }

    // Read the Excel file
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });

      // Get the first sheet from the workbook
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      // Assuming the Excel file contains only one row of data for this example
      const row = jsonData[0];
      setFormData({
        policyNumber: row.policyNumber,
        policyHolderName: row.policyHolderName,
        address: row.address,
        phone: row.phone,
        email: row.email,
        policyCost1: row.policyCost1,
        price1: row.price1,
        policyCost2: row.policyCost2,
        price2: row.price2,
        dueDate: row.dueDate,
      });
    };

    reader.readAsBinaryString(file);
  };

  const generateDocument = async () => {
    const {
      policyNumber,
      policyHolderName,
      address,
      phone,
      email,
      policyCost1,
      price1,
      policyCost2,
      price2,
      dueDate,
    } = formData;

    const totalCost = Number(price1) + Number(price2);

    try {
      // Fetch the template file
      const response = await fetch("/docs_template.docx");
      const templateBuffer = await response.arrayBuffer();

      // Load the template into PizZip
      const zip = new PizZip(templateBuffer);

      // Create a Docxtemplater instance
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Set the data for placeholders
      doc.setData({
        policyNumber,
        policyHolderName,
        address,
        phone,
        email,
        policyCost1,
        price1,
        policyCost2,
        price2,
        totalCost,
        dueDate,
      });

      // Render the document with the data
      doc.render();

      // Export the document as a blob
      const out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      saveAs(out, "GeneratedPolicyDocument.docx");
    } catch (error) {
      console.error("Error generating document:", error);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center py-8">
      <div className="max-w-4xl w-full bg-white rounded-lg shadow-lg p-6 space-y-6">
        <h1 className="text-3xl font-semibold text-center text-indigo-600">
          Policy Document Generator
        </h1>

        {/* File input for Excel upload */}
        <div className="flex justify-center">
          <div className="space-x-4">
            <input
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              className="file:py-2 file:px-4 file:bg-indigo-500 file:text-white file:rounded-md file:cursor-pointer hover:file:bg-indigo-600 transition"
            />
            <button
              onClick={handleFileUpload}
              className="px-6 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 transition"
            >
              Upload Excel File
            </button>
          </div>
        </div>

        {/* Show the data extracted from the Excel file */}
        <div className="bg-gray-100 p-4 rounded-md shadow-sm">
          <h3 className="text-lg font-medium">Uploaded Data:</h3>
          <pre className="text-sm text-gray-700">{JSON.stringify(formData, null, 2)}</pre>
        </div>

        <div className="flex justify-center">
          <button
            onClick={generateDocument}
            className="px-6 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition"
          >
            Generate Document
          </button>
        </div>
      </div>
    </div>
  );
};

export default App;
