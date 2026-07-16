import { BrowserRouter, Routes, Route } from "react-router-dom";
import Dashboard from "./pages/Dashboard";
import EmailPage from "./pages/EmailPage";
import QuotesPage from "./pages/QuotesPage";
import QuoteDetailPage from "./pages/QuoteDetailPage";
import SettingsPage from "./pages/SettingsPage";
import { QuoteJobsProvider } from "./context/QuoteJobsContext";
import BackgroundQuoteBar from "./components/BackgroundQuoteBar";

export default function App() {
  return (
    <BrowserRouter>
      <QuoteJobsProvider>
        <Routes>
          <Route path="/" element={<Dashboard />} />
          <Route path="/email" element={<EmailPage />} />
          <Route path="/quotes" element={<QuotesPage />} />
          <Route path="/quotes/:jobId" element={<QuoteDetailPage />} />
          <Route path="/settings" element={<SettingsPage />} />
        </Routes>
        <BackgroundQuoteBar />
      </QuoteJobsProvider>
    </BrowserRouter>
  );
}
