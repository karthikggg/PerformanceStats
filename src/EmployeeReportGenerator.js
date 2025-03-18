import React, { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
const EmployeeReportGenerator = () => {
 const [templateFile, setTemplateFile] = useState(null);
 const [newEmployeeId, setNewEmployeeId] = useState("");
 const [newEmployeeName, setNewEmployeeName] = useState("");
 const [employeeData, setEmployeeData] = useState([
    { id: 10409934, name: "Vinay UK" },
    { id: 10411916, name: "Dhruv R Doshi" },
    { id: 10411772, name: "Jenitha Inbarasi" },
    { id: 10411774, name: "Varsha Mahadevan" },
    { id: 10412819, name: "Saquib Tanweer" },
    { id: 10412814, name: "Jeff Rohit" },
    { id: 10412811, name: "Savitha Panneerselvan" },
    { id: 10412996, name: "Arjun Thirumalaikumar" },
    { id: 10415412, name: "Veera Sabarinathan" },
    { id: 10415413, name: "Hemarupa Karthikeyan" },
    { id: 10416036, name: "Bala Thirupathi Raaja" },
    { id: 10417208, name: "Pragadeeshwaran Ganesan" },
    { id: 10417209, name: "Akshay Kumar P" },
    { id: 10417367, name: "Shalini Subramanian" },
    { id: 10418289, name: "Sivasankari Arumugam" },
    { id: 10418626, name: "Prabhakaran Sekar" },
    { id: 10418869, name: "Santhosh Khanth" },
    { id: 10419645, name: "Mohammed Umar Mansoor" },
    { id: 10419682, name: "Irene Devakirubai" },
    { id: 10420136, name: "Sindhuja Prabakaran" },
    { id: 10420137, name: "Siddharthan Mayilsamy" },
    { id: 10420134, name: "Harishwar Nesamani" },
    { id: 10420135, name: "Kowsalya G" },
    { id: 10421099, name: "Anshuman Dey" },
    { id: 10421101, name: "Ajit Balaji" },
    { id: 10421103, name: "Kumaran Ramachandran" },
    { id: 10421093, name: "Sabariraj Iyyappan" },
    { id: 10421094, name: "Keerthana Ganesh" },
    { id: 10421096, name: "Sabarish Gupta Obilisetti" },
    { id: 10421097, name: "Rajamouli Ramaiyan" },
    { id: 10421667, name: "Thaanu Kumar M" },
    { id: 10421670, name: "Sounderajan Rengasamy" },
    { id: 10421775, name: "Gowtham Rathinasamy" },
    { id: 10421772, name: "Dinesh Kumar Saravanan" },
    { id: 10421934, name: "Sanjay kumar.M" },
    { id: 10422354, name: "Akash.M Murali" },
    { id: 10422355, name: "Shaik Afrid" },
    { id: 10422783, name: "Yazhini Krishnamoorthy" },
    { id: 10422784, name: "Sabana Satik" },
    { id: 10423034, name: "Allan Augustine" },
    { id: 10423370, name: "Ganesa Murugan" },
    { id: 10423978, name: "Vijayalakshmi Janakiraman" },
    { id: 10425117, name: "Akshay Gopakumar" },
    { id: 10425125, name: "Sachin Rajesh" },
    { id: 10425121, name: "Sathish Kumar Sankaranagappan" },
    { id: 10425123, name: "Manju Kolli" },
    { id: 10425118, name: "Sri Balaji Sudharshan" },
    { id: 10425115, name: "Parthasarathy Letchumanan" },
    { id: 10425124, name: "Siva Ganesh Santhanam" },
    { id: 10425415, name: "Jenithson Thommai" },
    { id: 10425416, name: "Saranesh Duraisamy" },
    { id: 10426079, name: "Rex Fleming" },
    { id: 10426890, name: "GN Karthik" },
    { id: 10426931, name: "Divakkar Varagunan" },
    { id: 10428929, name: "Mohammed Wihaj" },
    { id: 10429597, name: "Asmitha Gnanaprakash" },
    { id: 10428930, name: "Kaviprabha G" },
    { id: 10428931, name: "Gurumoorthy Vijayarangan" },
    { id: 10428932, name: "Rekha B" },
    { id: 10428935, name: "Krishna Chaitanya" },
    { id: 10428934, name: "Hariharan N" },
    { id: 10429583, name: "Gnana Jenifer Wilciya" },
    { id: 10429979, name: "Karthikeyan Shankar" },
    { id: 10430834, name: "Aishwarya Rajamohan" },
    { id: 10430832, name: "Vinod Ram" },
    { id: 10430830, name: "Deepika Raghuraj" },
    { id: 10431147, name: "Divya Dharshini" },
    { id: 10431142, name: "Melwin Manoj" },
    { id: 10431141, name: "Anusree Anil" },
    { id: 10433153, name: "Ramprakash Rajan" },
    { id: 10433155, name: "Bhargavi Baskaran" },
    { id: 10433152, name: "Ayyapparaj Dhamodhaan" },
    { id: 10433154, name: "Harshaavardhan Subramani" },
    { id: 10432953, name: "Pooja Raghavendra" },
    { id: 10433441, name: "Bhuvan Balasubramanian" },
    { id: 10445740, name: "Swetha Mani" },
    { id: 10446572, name: "Vijay Kumar R" },
    { id: 10446964, name: "Arul Mani Joseph" },
    { id: 10446962, name: "Manoj Rajasekaran" },
    { id: 10446965, name: "Ali Mehran Kandrikar" },
    { id: 10446967, name: "Naveen Srinivasan" },
  { id: 10446966, name: "Kishore Ganesan" },
  { id: 10447158, name: "Prasanth Rajendran" },
  { id: 10447160, name: "Avi Sharma" },
  { id: 10447662, name: "Veeravisvavinayagam Kumaravelu" },
  { id: 10447168, name: "Jeevanandam Ruthramurthy" },
  { id: 10447398, name: "Epsi Surendran" },
  { id: 10447277, name: "Saran Kumar G" },
  { id: 10447157, name: "Karthik Govindasamy" },
  { id: 10447281, name: "Tharun Kumar V" },
  { id: 10447280, name: "Nitish Kumar" },
  { id: 10447397, name: "Divya Barani Karthikeyean" },
  { id: 10447156, name: "Vishwa Alagiri" },
  { id: 10447155, name: "Shantha Kumar Saravanan" },
  { id: 10447163, name: "Meenakshi Maragathavel" },
  { id: 10447396, name: "Durairaj Saravanakumar" },
  { id: 10447663, name: "Sai Kumar C" },
  { id: 10447166, name: "Priyadharshini Mohan" },
  { id: 10447162, name: "Vishnu Bose" },
  { id: 10447273, name: "Lakshmi Aishwarya Ratakondala" },
  { id: 10447276, name: "Shanmuga Priya. Ramesh" },
  { id: 10447165, name: "Priyea Dharshani B" },
  { id: 10447164, name: "Yuvaraj Selvam" },
  { id: 10447275, name: "Ashwin Kumar S" },
  { id: 10447167, name: "Janani Venkatesalu" },
  { id: 10447161, name: "Jayasree Mohanakrishnan" },
  { id: 10447334, name: "Kiranraj Ravichandran" },
  { id: 10447664, name: "Vidhul Jothi Senthil Nathan" },
  { id: 10447335, name: "Priyadharshini James" },
  { id: 10447336, name: "Moneshwar Devaraj" },
  { id: 10447337, name: "Shifhana Banu Usain" },
  { id: 10447338, name: "Goutham Sakthivel" },
  { id: 10447279, name: "Dilip Suresh" },
  { id: 10447665, name: "Kishore Sivalingam" },
  { id: 10448179, name: "Dhuruva Gowshik Ganesan" },
  { id: 10429329, name: "Aarthi Madhan" },
  { id: 10429332, name: "Ajay Dhandapani" },
  { id: 10429336, name: "Akash Sampath" },
  { id: 10429339, name: "Arun Sajeev" },
  { id: 10429340, name: "Aswini Haribabu" },
  { id: 10429343, name: "Augustina Albert Sagayaraj" },
  { id: 10429346, name: "Deepika Subramani" },
  { id: 10429349, name: "Dhanalakshmi Sundar" },
  { id: 10429354, name: "Harihara Ponnaiah" },
  { id: 10429259, name: "Kamaleeshwari Sasi Kapoor Singh" },
  { id: 10429360, name: "Nithish Thivya" },
  { id: 10429367, name: "Praveen Kumar Thanigaiarasu" },
  { id: 10429361, name: "Rajeshwari Rajagopal" },
  { id: 10429368, name: "Yuvasree Balasubramaniam" },
  { id: 10429384, name: "Saran T" },
  { id: 10448387, name: "Karthick Gurunathan" },
  { id: 10448384, name: "Nishanthini Umapathy" },
  { id: 10448382, name: "Rohit Subramani" },
  { id: 10448381, name: "Samyuktha Balakrishnaian" },
  { id: 10448390, name: "Vijayalakshmi Dhanabalan" },
  { id: 10448377, name: "Vignesh Murugan" },
  { id: 10448376, name: "Mariya Antony Britto" },
  { id: 10448380, name: "Sarath Kumar Ravikumar" },
  { id: 10448379, name: "Surbash Lakshmi Gandhan" },
  { id: 10448386, name: "Karthikeyan Panchavaranam" },
  { id: 10448388, name: "Dharsini Nethaji" },
  { id: 10448378, name: "Tamilarasi Balamurugan" },
  { id: 10448385, name: "Nadhiya Siva Subramanian" },
  { id: 10448393, name: "Sathish Kumar Venkatesan" },
  { id: 10448959, name: "Balanaveena Arjunan" },
  { id: 10449931, name: "Angu selvam Murugan" },
  { id: 10450247, name: "Ranjana Mohan" },
  { id: 10450249, name: "Pradeep Joel Xavier" },
  { id: 10450402, name: "Vedhasree Manivannan" },
  { id: 10450312, name: "Balaji Ashok Kumar" },
  { id: 10451279, name: "Palani Raja Vellaisamy" },
  { id: 10451121, name: "Anitha Ananthan" },
  { id: 10451414, name: "Sarathirajan K" },
  { id: 10451357, name: "Naveen Manikandan" },
  { id: 10451358, name: "Siddhanth Ramesh" },
  { id: 10453089, name: "Divya Shree" },
  { id: 10453088, name: "Sneha Hari Doss" },
  { id: 10453090, name: "Manoj Thiruppathi" },
  { id: 10453092, name: "Sandhiya Kollapuri" },
  { id: 10453152, name: "Kirthika Jayaraman" },
  { id: 10457539, name: "Saranya Selvamani" },
  { id: 10466495, name: "Naveen Kumar Sankar" },
  { id: 10468964, name: "Hemavathy Rajendran" },
  { id: 10470269, name: "Amrutha Rajan" },
  { id: 10471150, name: "Nivedhaa Mohankumar" },
  { id: 10447088, name: "Tarpan Ghoshal" },
  { id: 10479182, name: "Anurag M" },
  { id: 10479183, name: "Uday Kumar" },
  { id: 10479181, name: "Sabari Ganesh K" },
  { id: 10480914, name: "Gowthami Jayashankar" },
  { id: 10481531, name: "Saran Kirthic" },
  { id: 10480915, name: "Bhavani Dhanabalan" },
  { id: 10480917, name: "Yugeshwaran Aroumougam" },
  { id: 10481530, name: "Sonia Selva Kumar" },
  { id: 10484450, name: "Mahalakshmi Nagaraj" },
  { id: 10480916, name: "Shayan Ahmed Viringipuram" },
  { id: 10488858, name: "Harini S K" },
  { id: 10508240, name: "Iswarya Jayabalan" },
  { id: 10470689, name: "Sudha Birendarkumar" },
  { id: 10470691, name: "Naveen Kumar Anandan" },
  { id: 10470693, name: "Priya Dharshini K" },
  { id: 10470993, name: "Ritesh Suresh" },
  { id: 10470976, name: "Deepika Sampath Kumar" },
  { id: 10470692, name: "Sruthi Mathivanan" },
  { id: 10471128, name: "Rangarajan Basker" },
  { id: 10471013, name: "Tarun Akash Pazhani S" },
  { id: 10470694, name: "Rojini.S Sathish Kumar" },
  { id: 10471007, name: "Akash N Natarajan C" },
  { id: 10470998, name: "Madhumitha.C Chandhiran.N" },
  { id: 10470997, name: "Najir Hussain Nashim Miyan" },
  { id: 10470679, name: "Logeshwari S Sundaramoorthy" },
  { id: 10514086, name: "Yashvanth Munusamy" },
  { id: 10514083, name: "Somalakshmi Dhanachezhiyan" },
  { id: 10514084, name: "Srilekha P" },
  { id: 10514076, name: "Pooja Gnanaprakasam" },
  { id: 10514077, name: "Dhanush Siva" },
  { id: 10514337, name: "Mohamed Jakeria" },
  { id: 10523035, name: "Lavanya Mahanti" },
  { id: 10524417, name: "Karthick Kumar" },
  { id: 10523034, name: "Krishnaraj Mohan" },
  { id: 10544112, name: "Keerthana J" },
  { id: 10544116, name: "Govarthan Mohan" },
  { id: 10544115, name: "Devakumar Y" },
  { id: 10544114, name: "Monisha Babu" },
  { id: 10544117, name: "Mahalakshmi G" },
  { id: 10544113, name: "Sathish E" },
 ]);
 // Handle template file upload
 const handleTemplateUpload = (event) => {
   const file = event.target.files[0];
   if (file) {
     setTemplateFile(file);
   }
 };
 // Handle adding a new employee
 const handleAddEmployee = () => {
   if (!newEmployeeId || !newEmployeeName) {
     alert("Please enter both Employee ID and Name.");
     return;
   }
   const newEmployee = {
     id: parseInt(newEmployeeId, 10),
     name: newEmployeeName.trim(),
   };
   setEmployeeData([...employeeData, newEmployee]);
   // Clear inputs after adding
   setNewEmployeeId("");
   setNewEmployeeName("");
 };
 // Generate employee reports
 const generateReports = async () => {
   if (!templateFile) {
     alert("Please upload a template file first.");
     return;
   }
   try {
     const reader = new FileReader();
     reader.onload = async (e) => {
       try {
         const arrayBuffer = e.target.result;
         const workbook = new ExcelJS.Workbook();
         await workbook.xlsx.load(arrayBuffer);
         if (workbook.worksheets.length === 0) {
           alert("Template file is empty or invalid.");
           return;
         }
         const templateSheet = workbook.worksheets[0];
         const newWorkbook = new ExcelJS.Workbook();
         employeeData.forEach((employee) => {
           const newSheet = newWorkbook.addWorksheet(employee.name);
           // Copy column properties
           templateSheet.columns.forEach((col, index) => {
             newSheet.getColumn(index + 1).width = col.width;
             newSheet.getColumn(index + 1).style = { ...col.style };
           });
           // Copy row & cell properties
           templateSheet.eachRow((row, rowNumber) => {
             const newRow = newSheet.getRow(rowNumber);
             row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
               const newCell = newRow.getCell(colNumber);
               newCell.value = cell.value;
               newCell.style = { ...cell.style };
             });
             newRow.commit();
           });
           // Set Employee ID in A1
           newSheet.getCell("A1").value = employee.id;
         });
         // Generate and download the Excel file
         const buffer = await newWorkbook.xlsx.writeBuffer();
         const blob = new Blob([buffer], {
           type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         });
         saveAs(blob, "Employee_Reports.xlsx");
       } catch (error) {
         console.error("Error processing Excel file:", error);
         alert("An error occurred while processing the file. Check the console for details.");
       }
     };
     reader.readAsArrayBuffer(templateFile);
   } catch (error) {
     console.error("File processing error:", error);
     alert("An unexpected error occurred. Check the console for details.");
   }
 };
 return (
<div className="container mt-5 p-4 border rounded shadow bg-light">
<h2 className="text-center fw-bold text-primary mb-4">New Stats Process</h2>
     {/* Template File Upload */}
<div className="mb-4">
<label htmlFor="template-upload" className="form-label">
        <h4> Upload Template File:</h4>
</label>
<input
         id="template-upload"
         type="file"
         accept=".xlsx"
         onChange={handleTemplateUpload}
         className="form-control"
       />
</div>
     {/* Add New Employee Section */}
<div className="mb-4">
<h3>Add New Employee</h3>
<div className="mb-3">
<label htmlFor="employee-id" className="form-label">
           Employee ID:
</label>
<input
           id="employee-id"
           type="text"
           value={newEmployeeId}
           onChange={(e) => setNewEmployeeId(e.target.value)}
           placeholder="Enter Employee ID"
           className="form-control"
         />
</div>
<div className="mb-3">
<label htmlFor="employee-name" className="form-label">
           Employee Name:
</label>
<input
           id="employee-name"
           type="text"
           value={newEmployeeName}
           onChange={(e) => setNewEmployeeName(e.target.value)}
           placeholder="Enter Employee Name"
           className="form-control"
         />
</div>
<button className="btn btn-primary" onClick={handleAddEmployee}>
         Add Employee
</button>
</div>
     {/* Generate Reports Button */}
<div className="mb-4">
<button className="btn btn-success" onClick={generateReports}>
         Generate Reports
</button>
</div>
</div>
 );
};
export default EmployeeReportGenerator;