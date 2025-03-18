import React from "react";
import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";
import "bootstrap/dist/css/bootstrap.min.css";
import StatsAutomate from "./StatsAutomate";
import EmployeeReportGenerator from "./EmployeeReportGenerator";
const App = () => {
 return (
<Router>
<div className="container mt-5">
<Routes>
         {/* Home Page with Two Buttons */}
<Route
           path="/"
           element={
<div className="text-center mt-5 p-4 border rounded shadow bg-light">
<h1 className="mb-4">Stats Automation Tools</h1>
<div className="d-flex justify-content-center gap-3">
<Link to="/stats-automate" className="btn btn-primary btn-lg">
                   Stats Automate
</Link>
<Link
                   to="/employee-report-generator"
                   className="btn btn-success btn-lg"
>
                   New Stats File Generator
</Link>
</div>
</div>
           }
         />
         {/* Stats Automate Page */}
<Route path="/stats-automate" element={<StatsAutomate />} />
         {/* Employee Report Generator Page */}
<Route path="/employee-report-generator" element={<EmployeeReportGenerator />} />
</Routes>
</div>
</Router>
 );
};
export default App;