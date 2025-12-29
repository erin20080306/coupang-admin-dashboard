import { Navigate, Route, Routes } from 'react-router-dom';
import { ProtectedRoute } from './components/ProtectedRoute';
import { isAuthed } from './lib/auth';
import AttendancePage from './pages/AttendancePage';
import DashboardPage from './pages/DashboardPage';
import LoginPage from './pages/LoginPage';
import './App.css';

export default function App() {
  return (
    <Routes>
      <Route
        path="/"
        element={
          <ProtectedRoute>
            <DashboardPage />
          </ProtectedRoute>
        }
      />
      <Route
        path="/attendance"
        element={
          <ProtectedRoute>
            <AttendancePage />
          </ProtectedRoute>
        }
      />
      <Route
        path="/login"
        element={isAuthed() ? <Navigate to="/" replace /> : <LoginPage />}
      />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}
