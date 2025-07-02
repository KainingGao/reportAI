import { BrowserRouter as Router, Routes, Route, Link, useLocation } from "react-router-dom";
import SafetyReportPage from "./SafetyReportPage";
import ImageResizer from "./ImageResizer";
import "./styles.css";

function Navigation() {
  const location = useLocation();
  
  return (
    <nav className="navigation">
      <div className="nav-container">
        <h2 className="nav-title">工具集合</h2>
        <div className="nav-links">
          <Link 
            to="/" 
            className={`nav-link ${location.pathname === "/" ? "active" : ""}`}
          >
            安全检查报告生成
          </Link>
          <Link 
            to="/image-resizer" 
            className={`nav-link ${location.pathname === "/image-resizer" ? "active" : ""}`}
          >
            文档图片尺寸调整
          </Link>
        </div>
      </div>
    </nav>
  );
}

export default function App() {
  return (
    <Router>
      <div className="app-container">
        <Navigation />
        <main className="main-content">
          <Routes>
            <Route path="/" element={<SafetyReportPage />} />
            <Route path="/image-resizer" element={<ImageResizer />} />
          </Routes>
        </main>
      </div>
    </Router>
  );
} 