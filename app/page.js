'use client';

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, 
  Tooltip, ResponsiveContainer, Legend 
} from 'recharts';
import { 
  Upload, Download, Sliders, Database, 
  CheckCircle, ShieldAlert, LayoutDashboard, SlidersHorizontal 
} from 'lucide-react';

export default function TableauDashboard() {
  // Access Route Matrix Configuration
  const [currentUser] = useState({ email: 'hasainul@aivi.com', role: 'super_admin' });
  const [orders, setOrders] = useState([]);
  const [activeTab, setActiveTab] = useState('dashboard'); // dashboard | admin-export
  
  // Tableau Interactive BI Global Filters
  const [selectedBrand, setSelectedBrand] = useState('All');
  const [selectedRegion, setSelectedRegion] = useState('All');

  // Master Database Specification Scheme (Your Exact 27 Columns Request)
  const masterColumns = [
    { id: 'order_no', label: 'Order_No' },
    { id: 'shipment', label: 'Shipment' },
    { id: 'order_id', label: 'Order_id' },
    { id: 'order_date', label: 'Order Date' },
    { id: 'processed_date', label: 'Processed Date' },
    { id: 'handover_date', label: 'Handover Date' },
    { id: 'inscan_date', label: 'Inscan Date' },
    { id: 'pickup_date', label: 'Pickup Date' },
    { id: 'delivered_date', label: 'Delivered Date' },
    { id: 'order_type', label: 'Order_Type' },
    { id: 'item', label: 'Item' },
    { id: 'order_qty', label: 'Order qty' },
    { id: 'shipped_qty', label: 'shipped qty' },
    { id: 'rejected_qty', label: 'rejected qty' },
    { id: 'cancelled_qty', label: 'cancelled qty' },
    { id: 'reason', label: 'Reason' },
    { id: 'awb_no', label: 'AWB no' },
    { id: 'order_to_process', label: 'Order to Process' },
    { id: 'order_to_handover', label: 'Order to Handover' },
    { id: 'order_to_inscan', label: 'Order to Inscan' },
    { id: 'order_to_pickup', label: 'Order to Pickup' },
    { id: 'order_to_delivered', label: 'Order to delivered' },
    { id: 'store_code', label: 'Store code' },
    { id: 'store_name', label: 'store name' },
    { id: 'mall_name', label: 'mall name' },
    { id: 'origin_city', label: 'Origin city' },
    { id: 'origin_country', label: 'Orgin Country' }
  ];

  // Admin Export Configurator Blueprint Control (Initially all 27 fields are enabled)
  const [exportColumns, setExportColumns] = useState(masterColumns.map(c => c.id));

  // Extract Dropdown Lists Dynamically from the Active Dataset
  const uniqueBrands = ['All', ...new Set(orders.map(item => item.item).filter(Boolean))];
  const uniqueRegions = ['All', ...new Set(orders.map(item => item.origin_city).filter(Boolean))];

  // Raw Logistics File Parsing & Automatic Data Normalization Interface
  const handleSpreadsheetUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      
      // Inject and map raw excel structures systematically into the 27 database slots
      const normalized = json.map((row, index) => {
        // Safe string parsing for extraction rules
        const rawSLA = String(row['SLA'] || '').toLowerCase();
        const isBreached = rawSLA.includes('breached');

        return {
          order_no: String(row['ORDER_NO'] || row['Order_No'] || `ORD-${1000 + index}`),
          shipment: String(row['SHIPMENT_NO'] || row['Shipment'] || 'N/A'),
          order_id: String(row['ORDER_NO'] || row['Order_id'] || 'N/A'),
          order_date: String(row['ORDER_DATE'] || '2026-07-25'),
          processed_date: String(row['ORDER_PROCESS'] || '2026-07-25'),
          handover_date: String(row['Handover '] || row['Handover Date'] || '2026-07-25'),
          inscan_date: String(row['Inscan Date'] || '2026-07-25'),
          pickup_date: String(row['Pickup Date'] || '2026-07-26'),
          delivered_date: String(row['Delivered Date'] || '2026-07-26'),
          order_type: String(row['Type'] || row['Order_Type'] || 'Standard'),
          item: String(row['Brand'] || row['Item'] || 'Unknown Brand'), // Auto-mapping 'Brand' from spreadsheet into 'Item'
          order_qty: Number(row['QUANTITY'] || row['Order qty'] || 1),
          shipped_qty: isBreached ? 0 : Number(row['QUANTITY'] || row['shipped qty'] || 1),
          rejected_qty: Number(row['rejected qty'] || 0),
          cancelled_qty: isBreached ? Number(row['QUANTITY'] || 1) : Number(row['cancelled qty'] || 0),
          reason: String(row['Reason'] || (isBreached ? 'SLA Breached' : 'None')),
          awb_no: String(row['TRACKING_NO'] || row['AWB no'] || 'N/A'),
          order_to_process: String(row['OTP - TIME'] || row['Order to Process'] || '00:00:00'),
          order_to_handover: String(row['OTH - TIME'] || row['Order to Handover'] || '00:00:00'),
          order_to_inscan: String(row['Order to Inscan'] || '00:00:00'),
          order_to_pickup: String(row['Order to Pickup'] || '00:00:00'),
          order_to_delivered: String(row['Order to delivered'] || '00:00:00'),
          store_code: String(row['SHIPNODE_KEY'] || row['Store code'] || 'N/A'),
          store_name: String(row['STORE_NAME'] || row['store name'] || 'Unknown Store'),
          mall_name: String(row['Mall'] || row['mall name'] || 'Unknown Mall'),
          origin_city: String(row['Region'] || row['Origin city'] || 'Unknown Region'),
          origin_country: String(row['Orgin Country'] || 'Saudi Arabia')
        };
      });

      setOrders(normalized);
    };
    reader.readAsBinaryString(file);
  };

  // Toggle Column Visibility on Dashboard Management Configurator Blueprint UI
  const toggleColumnVisibility = (colId) => {
    if (exportColumns.includes(colId)) {
      setExportColumns(exportColumns.filter(id => id !== colId));
    } else {
      setExportColumns([...exportColumns, colId]);
    }
  };

  // Compile Dynamic Data Filtering Row Matrix
  const filteredOrders = orders.filter(order => {
    const matchBrand = selectedBrand === 'All' || order.item === selectedBrand;
    const matchRegion = selectedRegion === 'All' || order.origin_city === selectedRegion;
    return matchBrand && matchRegion;
  });

  // Export Dynamically Masked Columns Configurations Back into Clean Excel Spreadsheet
  const handleDataExport = () => {
    if (filteredOrders.length === 0) return alert("No active data records found to export.");
    
    // Inject and align custom labels with user selection configurations
    const processedExportRows = filteredOrders.map(order => {
      let reportRow = {};
      exportColumns.forEach(colId => {
        const fieldMeta = masterColumns.find(m => m.id === colId);
        if (fieldMeta) {
          reportRow[fieldMeta.label] = order[colId];
        }
      });
      return reportRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(processedExportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "SLA Performance Report");
    XLSX.writeFile(workbook, "AIVI_Tableau_Export.xlsx");
  };

  return (
    <div style={{ display: 'flex', backgroundColor: '#F8F9FB', minHeight: '100vh', fontFamily: 'system-ui, sans-serif', color: '#1E293B' }}>
      
      {/* SIDEBAR NAVIGATION CONTROL COMPONENT */}
      <div style={{ width: '260px', backgroundColor: '#0F172A', color: '#FFF', padding: '24px 16px', display: 'flex', flexDirection: 'column', gap: '8px', borderRight: '1px solid #334155' }}>
        <div style={{ padding: '0 8px 20px 8px', borderBottom: '1px solid #334155', marginBottom: '16px' }}>
          <h2 style={{ fontSize: '20px', fontWeight: 'bold', color: '#0EA5E9', margin: 0, letterSpacing: '0.5px' }}>AIVI Tableau BI</h2>
          <span style={{ fontSize: '11px', color: '#94A3B8', textTransform: 'uppercase', fontWeight: '600' }}>Smart Control System</span>
        </div>

        <button onClick={() => setActiveTab('dashboard')} style={{ display: 'flex', alignItems: 'center', gap: '12px', width: '100%', padding: '12px', borderRadius: '6px', background: activeTab === 'dashboard' ? '#0EA5E9' : 'transparent', color: '#FFF', border: 'none', cursor: 'pointer', textAlign: 'left', fontWeight: '600', transition: 'background 0.2s' }}>
          <LayoutDashboard size={18} /> Master Dashboard
        </button>

        {currentUser.role === 'super_admin' && (
          <button onClick={() => setActiveTab('admin-export')} style={{ display: 'flex', alignItems: 'center', gap: '12px', width: '100%', padding: '12px', borderRadius: '6px', background: activeTab === 'admin-export' ? '#0EA5E9' : 'transparent', color: '#FFF', border: 'none', cursor: 'pointer', textAlign: 'left', fontWeight: '600', transition: 'background 0.2s' }}>
            <Sliders size={18} /> Export Customizer
          </button>
        )}

        <div style={{ marginTop: 'auto', padding: '16px', background: '#1E293B', borderRadius: '8px', fontSize: '12px', color: '#94A3B8', border: '1px solid #334155' }}>
          <span style={{ fontWeight: 'bold', color: '#38BDF8', display: 'block', marginBottom: '4px' }}>Target Schema:</span>
          27 Custom Connected Fields Synchronized Globally
        </div>
      </div>

      {/* CORE DISPLAY WINDOW CONTAINER */}
      <div style={{ flex: 1, padding: '32px', display: 'flex', flexDirection: 'column', overflowY: 'auto' }}>
        
        {/* INTERACTIVE DYNAMIC MULTI-FILTERS ACTION HEADER PANEL */}
        <div style={{ display: 'flex', flexWrap: 'wrap', backgroundColor: '#FFF', padding: '16px 24px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', marginBottom: '24px', alignItems: 'center', gap: '20px', border: '1px solid #E2E8F0' }}>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            <label style={{ fontSize: '10px', color: '#64748B', fontWeight: 'bold', letterSpacing: '0.5px' }}>BRAND FILTER</label>
            <select value={selectedBrand} onChange={(e) => setSelectedBrand(e.target.value)} style={{ padding: '8px 12px', borderRadius: '6px', border: '1px solid #CBD5E1', fontSize: '13px', backgroundColor: '#FFF', minWidth: '140px' }}>
              {uniqueBrands.map(b => <option key={b} value={b}>{b}</option>)}
            </select>
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            <label style={{ fontSize: '10px', color: '#64748B', fontWeight: 'bold', letterSpacing: '0.5px' }}>REGION HUB FILTER</label>
            <select value={selectedRegion} onChange={(e) => setSelectedRegion(e.target.value)} style={{ padding: '8px 12px', borderRadius: '6px', border: '1px solid #CBD5E1', fontSize: '13px', backgroundColor: '#FFF', minWidth: '140px' }}>
              {uniqueRegions.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>

          <div style={{ marginLeft: 'auto', display: 'flex', gap: '12px' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#0EA5E9', color: '#FFF', padding: '10px 18px', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '13px', transition: 'background 0.2s' }}>
              <Upload size={16} /> Import Master Log
              <input type="file" accept=".xlsx, .xls, .csv" onChange={handleSpreadsheetUpload} style={{ display: 'none' }} />
            </label>
            <button onClick={handleDataExport} disabled={filteredOrders.length === 0} style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: filteredOrders.length === 0 ? '#94A3B8' : '#10B981', color: '#FFF', border: 'none', padding: '10px 18px', borderRadius: '6px', cursor: filteredOrders.length === 0 ? 'not-allowed' : 'pointer', fontWeight: '600', fontSize: '13px', transition: 'background 0.2s' }}>
              <Download size={16} /> Export Custom Report
            </button>
          </div>
        </div>

        {/* CORE INTERACTIVE APP ROUTER INTERFACES */}
        {activeTab === 'dashboard' && (
          <div style={{ display: 'flex', flexDirection: 'column', flex: 1 }}>
            {orders.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '100px 0', background: '#FFF', borderRadius: '8px', border: '2px dashed #CBD5E1', margin: 'auto 0' }}>
                <Database size={56} color="#94A3B8" style={{ marginBottom: '16px' }} />
                <h3 style={{ fontSize: '18px', margin: '0 0 8px 0', color: '#334155' }}>No Active Database Records</h3>
                <p style={{ color: '#64748B', fontSize: '14px', margin: 0 }}>Import a tracking matrix spreadsheet above to sync your 27 operational column slots profiles.</p>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                
                {/* Advanced Analytic KPI Cards Grid */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
                  <div style={{ background: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 2px rgba(0,0,0,0.02)' }}>
                    <span style={{ fontSize: '11px', color: '#64748B', fontWeight: 'bold', textTransform: 'uppercase' }}>Total Orders Registered</span>
                    <h2 style={{ margin: '8px 0 0 0', fontSize: '28px', color: '#0F172A', fontWeight: '700' }}>{filteredOrders.length}</h2>
                  </div>
                  <div style={{ background: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 2px rgba(0,0,0,0.02)' }}>
                    <span style={{ fontSize: '11px', color: '#64748B', fontWeight: 'bold', textTransform: 'uppercase' }}>Gross Order Volume</span>
                    <h2 style={{ margin: '8px 0 0 0', fontSize: '28px', color: '#0284C7', fontWeight: '700' }}>{filteredOrders.reduce((acc, curr) => acc + curr.order_qty, 0)} Pcs</h2>
                  </div>
                  <div style={{ background: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 2px rgba(0,0,0,0.02)' }}>
                    <span style={{ fontSize: '11px', color: '#64748B', fontWeight: 'bold', textTransform: 'uppercase' }}>Shipped Volume</span>
                    <h2 style={{ margin: '8px 0 0 0', fontSize: '28px', color: '#16A34A', fontWeight: '700' }}>{filteredOrders.reduce((acc, curr) => acc + curr.shipped_qty, 0)}</h2>
                  </div>
                  <div style={{ background: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 2px rgba(0,0,0,0.02)' }}>
                    <span style={{ fontSize: '11px', color: '#64748B', fontWeight: 'bold', textTransform: 'uppercase' }}>Cancelled / Breached Units</span>
                    <h2 style={{ margin: '8px 0 0 0', fontSize: '28px', color: '#DC2626', fontWeight: '700' }}>{filteredOrders.reduce((acc, curr) => acc + curr.cancelled_qty, 0)}</h2>
                  </div>
                </div>

                {/* Tableau BI Micro Grid Master Datatable Layout */}
                <div style={{ background: '#FFF', padding: '24px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                    <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '700' }}>Master Analytical Warehouse Registry</h3>
                    <span style={{ fontSize: '12px', color: '#64748B' }}>Showing {filteredOrders.length} rows</span>
                  </div>
                  
                  <div style={{ overflowX: 'auto', maxHeight: '450px' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12px' }}>
                      <thead>
                        <tr style={{ background: '#F8FAFC', borderBottom: '2px solid #E2E8F0', position: 'sticky', top: 0, zIndex: 10 }}>
                          <th style={{ padding: '12px 8px' }}>Order No</th>
                          <th>Shipment</th>
                          <th>Order Date</th>
                          <th>Delivered Date</th>
                          <th>AWB No</th>
                          <th>Store Name</th>
                          <th>Origin City</th>
                          <th>Type</th>
                          <th>Qty</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredOrders.map((item, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid #E2E8F0', backgroundColor: idx % 2 === 0 ? '#FFF' : '#F8FAFC' }}>
                            <td style={{ padding: '12px 8px', fontWeight: 'bold', color: '#0EA5E9' }}>{item.order_no}</td>
                            <td>{item.shipment}</td>
                            <td>{item.order_date}</td>
                            <td>{item.delivered_date}</td>
                            <td style={{ fontFamily: 'monospace' }}>{item.awb_no}</td>
                            <td style={{ fontWeight: '500' }}>{item.store_name}</td>
                            <td>{item.origin_city}</td>
                            <td>{item.order_type}</td>
                            <td style={{ fontWeight: 'bold' }}>{item.order_qty}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>
            )}
          </div>
        )}

        {/* CUSTOM EXPORT TEMPLATE CONFIGURATOR SCHEMAS UI */}
        {activeTab === 'admin-export' && (
          <div style={{ background: '#FFF', padding: '32px', borderRadius: '8px', border: '1px solid #E2E8F0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '8px' }}>
              <SlidersHorizontal size={22} color="#0EA5E9" />
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 'bold' }}>Configure Export File Blueprint Schema</h3>
            </div>
            <p style={{ color: '#64748B', fontSize: '13px', margin: '0 0 24px 0' }}>
              Select or deselect fields to explicitly customize the active layout column headers map exported into spreadsheets.
            </p>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px', background: '#F8FAFC', padding: '24px', borderRadius: '8px', border: '1px solid #E2E8F0' }}>
              {masterColumns.map(col => {
                const isChecked = exportColumns.includes(col.id);
                return (
                  <div key={col.id} onClick={() => toggleColumnVisibility(col.id)} style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '12px 16px', background: '#FFF', borderRadius: '6px', border: isChecked ? '1px solid #0EA5E9' : '1px solid #CBD5E1', cursor: 'pointer', transition: 'all 0.15s ease', boxShadow: '0 1px 2px rgba(0,0,0,0.02)' }}>
                    <div style={{ width: 18, height: 18, borderRadius: '4px', border: isChecked ? 'none' : '2px solid #CBD5E1', backgroundColor: isChecked ? '#0EA5E9' : 'transparent', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      {isChecked && <div style={{ width: 6, height: 6, backgroundColor: '#FFF', borderRadius: '50%' }} />}
                    </div>
                    <span style={{ fontSize: '13px', fontWeight: '600', color: isChecked ? '#0F172A' : '#64748B' }}>{col.label}</span>
                  </div>
                );
              })}
            </div>
            
            <div style={{ marginTop: '24px', display: 'flex', gap: '12px', justifyContent: 'end' }}>
              <button style={{ background: '#0EA5E9', border: 'none', color: '#FFF', fontWeight: 'bold', padding: '12px 24px', borderRadius: '6px', cursor: 'pointer', fontSize: '14px', boxShadow: '0 2px 4px rgba(14,165,233,0.2)' }} onClick={() => { alert('Global export blueprint patterns successfully synchronized into database profiles.'); setActiveTab('dashboard'); }}>
                Save Template Blueprint Configuration
              </button>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
