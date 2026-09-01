import React from 'react';
import { NavLink, Outlet } from 'react-router-dom';
import { FolderUp, Table2 } from 'lucide-react';

export const Layout: React.FC = () => {
  const tabClass = ({ isActive }: { isActive: boolean }) =>
    `flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium rounded-md transition-colors ${
      isActive ? 'bg-primary text-white' : 'text-gray-600 hover:bg-gray-100'
    }`;

  return (
    <div className="min-h-screen bg-gray-50">
      <nav className="bg-white border-b border-gray-200 px-4 py-2 flex items-center gap-2">
        <NavLink to="/" end className={tabClass}>
          <FolderUp className="w-4 h-4" />Converter
        </NavLink>
        <NavLink to="/editor" className={tabClass}>
          <Table2 className="w-4 h-4" />Manual Editor
        </NavLink>
      </nav>
      <Outlet />
    </div>
  );
};
