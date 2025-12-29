import { Navigate } from 'react-router-dom';
import { isAuthed } from '../lib/auth';

export function ProtectedRoute({ children }: { children: React.ReactNode }) {
  if (!isAuthed()) {
    return <Navigate to="/login" replace />;
  }
  return children;
}
