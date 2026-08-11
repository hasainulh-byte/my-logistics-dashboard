'use client';

import React, { useState } from 'react';
import { Shield, Users, Key, Save, ArrowLeft, CheckCircle2 } from 'lucide-react';
import Link from 'next/link';

export default function AdminConsole() {
  const [users, setUsers] = useState([
    { id: 1, name: 'Hasainul Haider', email: 'admin@aivi-ops.com', role: 'Super Admin', status: 'Active' },
    { id: 2, name: 'Logistics Analyst', email: 'analyst@aivi-ops.com', role: 'Viewer', status: 'Active' }
  ]);

  const [selectedRole, setSelectedRole] = useState('Super Admin');
  const [newEmail, setNewEmail] = useState('');
  const [saved, setSaved] = useState(false);

  const handleAddUser = (e) => {
    e.preventDefault();
    if (!newEmail) return;
    setUsers([...users, { id: Date.now(), name: newEmail.split('@')[0], email: newEmail, role: selectedRole, status: 'Active' }]);
    setNewEmail('');
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  };

  return (
    <div style={{ backgroundColor: '#FDFBF7', minHeight: '100vh', padding: '32px', fontFamily: 'sans-serif', color: '#333' }}>
      
      {/* HEADER */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '28px', borderBottom: '1px solid #EAE6DF', paddingBottom: '16px' }}>
        <div>
          <Link href="/" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', color: '#8B4513', textDecoration: 'none', fontSize: '13px', fontWeight: 'bold', marginBottom: '8px' }}>
            <ArrowLeft size={16} /> Back to Dashboard
          </Link>
          <h1 style={{ fontSize: '24px', fontWeight: 'bold', color: '#111', margin: 0 }}>AIVI-OPS Admin Control Panel</h1>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: '#FEF3C7', color: '#92400E', padding: '6px 12px', borderRadius: '20px', fontSize: '12px', fontWeight: 'bold' }}>
          <Shield size={16} /> Super Admin Access Enabled
        </div>
      </div>

      {saved && (
        <div style={{ background: '#D1FAE5', color: '#065F46', padding: '12px', borderRadius: '6px', marginBottom: '20px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '14px' }}>
          <CheckCircle2 size={18} /> User access rights updated successfully!
        </div>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '24px' }}>
        
        {/* ADD USER CARD */}
        <div style={{ background: '#FFF', padding: '24px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
          <h3 style={{ fontSize: '16px', fontWeight: 'bold', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Key size={18} color="#8B4513" /> Provision User Access
          </h3>
          <form onSubmit={handleAddUser}>
            <div style={{ marginBottom: '14px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', marginBottom: '4px', color: '#666' }}>User Email</label>
              <input 
                type="email" 
                required 
                placeholder="user@domain.com" 
                value={newEmail} 
                onChange={(e) => setNewEmail(e.target.value)} 
                style={{ width: '100%', padding: '8px', border: '1px solid #CCC', borderRadius: '4px', fontSize: '13px' }}
              />
            </div>
            <div style={{ marginBottom: '20px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 'bold', marginBottom: '4px', color: '#666' }}>Role Assignment</label>
              <select value={selectedRole} onChange={(e) => setSelectedRole(e.target.value)} style={{ width: '100%', padding: '8px', border: '1px solid #CCC', borderRadius: '4px', fontSize: '13px' }}>
                <option value="Super Admin">Super Admin (Full Access + Export + Column Customizer)</option>
                <option value="Control Tower Manager">Control Tower Manager (View + Export)</option>
                <option value="Viewer">Viewer (Read-Only KPI & Data View)</option>
              </select>
            </div>
            <button type="submit" style={{ width: '100%', background: '#8B4513', color: '#FFF', border: 'none', padding: '10px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '8px' }}>
              <Save size={16} /> Save User Permissions
            </button>
          </form>
        </div>

        {/* ACTIVE USERS MATRIX LIST */}
        <div style={{ background: '#FFF', padding: '24px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
          <h3 style={{ fontSize: '16px', fontWeight: 'bold', marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Users size={18} color="#8B4513" /> Active Platform Users
          </h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', textAlign: 'left' }}>
            <thead>
              <tr style={{ borderBottom: '2px solid #EAE6DF', color: '#666', background: '#F9FAFB' }}>
                <th style={{ padding: '10px' }}>Name / Email</th>
                <th>Assigned Role</th>
                <th>Status</th>
                <th style={{ textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {users.map(u => (
                <tr key={u.id} style={{ borderBottom: '1px solid #F3F4F6' }}>
                  <td style={{ padding: '12px 10px', fontWeight: '500' }}>
                    {u.name}<br/>
                    <span style={{ fontSize: '11px', color: '#999' }}>{u.email}</span>
                  </td>
                  <td>
                    <span style={{ 
                      padding: '4px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold',
                      background: u.role === 'Super Admin' ? '#FEF3C7' : '#E0F2FE',
                      color: u.role === 'Super Admin' ? '#92400E' : '#0369A1'
                    }}>
                      {u.role}
                    </span>
                  </td>
                  <td>
                    <span style={{ color: '#10B981', fontWeight: 'bold', fontSize: '12px' }}>● {u.status}</span>
                  </td>
                  <td style={{ textAlign: 'right' }}>
                    <button style={{ color: '#DC2626', border: 'none', background: 'transparent', cursor: 'pointer', fontSize: '12px', fontWeight: 'bold' }}>Revoke</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

      </div>
    </div>
  );
}
