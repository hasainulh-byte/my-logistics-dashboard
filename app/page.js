'use client';

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, 
  Tooltip, ResponsiveContainer, BarChart
} from 'recharts';
import { Upload, Download, SlidersHorizontal, CheckCircle, LayoutDashboard, Table, Search } from 'lucide-react';

export default function AiviOpsDashboard() {
  // Navigation State (Tabs)
  const [activeTab, setActiveTab] = useState('kpi'); // 'kpi' | 'data'

  // Core Operational Data States
  const [orders, setOrders] = useState([]);
  const [showExportCustomizer, setShowExportCustomizer] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  
  // Dynamic Global Filter States
  const [dateRange, setDateRange] = useState({ start: '2026-07-25', end: '2026-07-31' });
  const [selectedBrand, setSelectedBrand] = useState('All');
  const [selectedRegion, setSelectedRegion] = useState('All');
  const [selectedType, setSelectedType] = useState('All');

  // Master Database Specification Schema (Your 27 Columns Matrix)
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

  // Admin Export Configurator Blueprint Control
  const [exportColumns, setExportColumns] = useState(masterColumns.map(c => c.id));

  // Dynamic dropdown lists from active dataset
  const uniqueBrands = ['All', ...new Set(orders.map(item => item.item).filter(Boolean))];
  const uniqueRegions = ['All', ...new Set(orders.map(item => item.origin_city).filter(Boolean))];
  const uniqueTypes = ['All', ...new Set(orders.map(item => item.order_type).filter(Boolean))];

  // Excel File Upload & Normalization Logic
  const handleSpreadsheetUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      
      const normalized = json.map((row, index) => {
        const rawSLA = String(row['SLA'] || '').toLowerCase();
        const isBreached = rawSLA.includes('breached');

        return {
          order_no: String(row['ORDER_NO'] || row['Order_No'] || `ORD-${1000 + index}`),
          shipment: String(row['SHIPMENT_NO'] || row['Shipment'] || 'N/A'),
          order_id: String(row['ORDER_NO'] || row['Order_id'] || 'N/A'),
          order_date: String(row['ORDER_DATE'] || '2026-07-25').split(' ')[0],
          processed_date: String(row['ORDER_PROCESS'] || '2026-07-25'),
          handover_date: String(row['Handover '] || row['Handover Date'] || '2026-07-25'),
          inscan_date: String(row['Inscan Date'] || '2026-07-25'),
          pickup_date: String(row['Pickup Date'] || '2026-07-26'),
          delivered_date: String(row['Delivered Date'] || '2026-07-26'),
          order_type: String(row['Type'] || row['Order_Type'] || 'Standard'),
          item: String(row['Brand'] || row['Item'] || 'Unknown'), 
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
          origin_city: String(row['Region'] || row['Origin city'] || 'Unknown Hub'),
          origin_country: String(row['Orgin Country'] || 'Saudi Arabia')
        };
      });
      setOrders(normalized);
    };
    reader.readAsBinaryString(file);
  };

  // Filter Active View Rows
  const filteredOrders = orders.filter(order => {
    const matchBrand = selectedBrand === 'All' || order.item === selectedBrand;
    const matchRegion = selectedRegion === 'All' || order.origin_city === selectedRegion;
    const matchType = selectedType === 'All' || order.order_type === selectedType;
    const matchSearch = searchTerm === '' || 
      order.order_no.toLowerCase().includes(searchTerm.toLowerCase()) ||
      order.store_name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      order.awb_no.toLowerCase().includes(searchTerm.toLowerCase());
    return matchBrand && matchRegion && matchType && matchSearch;
  });

  // Aggregations for Chart Data Components
  const getDailyMetrics = () => {
    const dailyMap = {};
    filteredOrders.forEach(o => {
      if (!dailyMap[o.order_date]) dailyMap[o.order_date] = { date: o.order_date, Orders: 0, Ontime: 0 };
      dailyMap[o.order_date].Orders++;
      if (o.cancelled_qty === 0) dailyMap[o.order_date].Ontime++;
    });
    return Object.values(dailyMap).map(d => ({
      ...d,
      "Ontime %": Math.round((d.Ontime / d.Orders) * 100)
    })).sort((a,b) => new Date(a.date) - new Date(b.date));
  };

  const getCategoricalMetrics = (key) => {
    const map = {};
    filteredOrders.forEach(o => {
      const val = o[key] || 'Unknown';
      if (!map[val]) map[val] = { name: val, Orders: 0, Ontime: 0 };
      map[val].Orders++;
      if (o.cancelled_qty === 0) map[val].Ontime++;
    });
    return Object.values(map).map(item => ({
      ...item,
      "Ontime %": Math.round((item.Ontime / item.Orders) * 100)
    }));
  };

  const getStoreRanking = (top = true) => {
    const storeMap = {};
    filteredOrders.forEach(o => {
      if (!storeMap[o.store_name]) storeMap[o.store_name] = { name: o.store_name, brand: o.item, region: o.origin_city, Orders: 0, Ontime: 0, Breached: 0 };
      storeMap[o.store_name].Orders++;
      if (o.cancelled_qty === 0) storeMap[o.store_name].Ontime++;
      else storeMap[o.store_name].Breached++;
    });
    const metrics = Object.values(storeMap).map(s => ({
      ...s,
      rate: Math.round((s.Ontime / s.Orders) * 100)
    }));
    return top ? metrics.sort((a,b) => b.rate - a.rate).slice(0, 10) : metrics.sort((a,b) => a.rate - b.rate).slice(0, 10);
  };

  // Customized Export Function
  const handleDataExport = () => {
    if (filteredOrders.length === 0) return alert("No operational rows available to export.");
    const reportData = filteredOrders.map(order => {
      let row = {};
      exportColumns.forEach(id => {
        const meta = masterColumns.find(m => m.id === id);
        if (meta) row[meta.label] = order[id];
      });
      return row;
    });
    const worksheet = XLSX.utils.json_to_sheet(reportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Customized Report");
    XLSX.writeFile(workbook, "AIVI_OPS_Export.xlsx");
  };

  return (
    <div style={{ backgroundColor: '#FDFBF7', minHeight: '100vh', padding: '24px', fontFamily: 'sans-serif', color: '#333' }}>
      
      {/* MAIN HEADER SECTION */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #EAE6DF', paddingBottom: '16px', marginBottom: '20px' }}>
        <div>
          <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#888', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Fulfillment · SLA · Store Performance</span>
          <h1 style={{ fontSize: '26px', fontWeight: 'bold', color: '#111', marginTop: '2px' }}>AIVI-OPS Dashboard</h1>
        </div>
        
        {/* ACTION BUTTONS */}
        <div style={{ display: 'flex', gap: '10px' }}>
          <button onClick={() => setShowExportCustomizer(!showExportCustomizer)} style={{ display: 'flex', alignItems: 'center', gap: '6px', border: '1px solid #CCC', padding: '8px 14px', borderRadius: '6px', backgroundColor: '#FFF', cursor: 'pointer', fontSize: '13px', fontWeight: '600' }}>
            <SlidersHorizontal size={14} /> Column Customizer
          </button>
          <label style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#8B4513', color: '#FFF', padding: '8px 14px', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '13px' }}>
            <Upload size={14} /> Import Log Sheet
            <input type="file" accept=".xlsx, .xls, .csv" onChange={handleSpreadsheetUpload} style={{ display: 'none' }} />
          </label>
          <button onClick={handleDataExport} disabled={filteredOrders.length === 0} style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#10B981', color: '#FFF', border: 'none', padding: '8px 14px', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '13px' }}>
            <Download size={14} /> Export File
          </button>
        </div>
      </div>

      {/* TAB NAVIGATION SWITCHER BAR */}
      <div style={{ display: 'flex', gap: '12px', marginBottom: '20px', borderBottom: '2px solid #EAE6DF', paddingBottom: '4px' }}>
        <button 
          onClick={() => setActiveTab('kpi')} 
          style={{ 
            display: 'flex', alignItems: 'center', gap: '8px', 
            padding: '10px 20px', border: 'none', background: 'transparent', 
            borderBottom: activeTab === 'kpi' ? '3px solid #8B4513' : '3px solid transparent', 
            color: activeTab === 'kpi' ? '#8B4513' : '#666', 
            fontWeight: 'bold', fontSize: '14px', cursor: 'pointer' 
          }}
        >
          <LayoutDashboard size={18} /> KPI Performance
        </button>

        <button 
          onClick={() => setActiveTab('data')} 
          style={{ 
            display: 'flex', alignItems: 'center', gap: '8px', 
            padding: '10px 20px', border: 'none', background: 'transparent', 
            borderBottom: activeTab === 'data' ? '3px solid #8B4513' : '3px solid transparent', 
            color: activeTab === 'data' ? '#8B4513' : '#666', 
            fontWeight: 'bold', fontSize: '14px', cursor: 'pointer' 
          }}
        >
          <Table size={18} /> Data View ({filteredOrders.length} Rows)
        </button>
      </div>

      {/* EXPORT CUSTOMIZER OVERLAY */}
      {showExportCustomizer && (
        <div style={{ background: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF', marginBottom: '20px' }}>
          <h4 style={{ margin: '0 0 10px 0', fontSize: '14px', fontWeight: 'bold' }}>Select Allowed Fields Blueprint for Excel Export Profile</h4>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', background: '#FAFAFA', padding: '12px', borderRadius: '6px' }}>
            {masterColumns.map(c => {
              const checked = exportColumns.includes(c.id);
              return (
                <div key={c.id} onClick={() => setExportColumns(checked ? exportColumns.filter(id => id !== c.id) : [...exportColumns, c.id])} style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '8px', background: '#FFF', borderRadius: '4px', border: checked ? '1px solid #8B4513' : '1px solid #DDD', cursor: 'pointer', fontSize: '12px' }}>
                  <CheckCircle size={14} color={checked ? '#8B4513' : '#CCC'} />
                  <span>{c.label}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* GLOBAL FILTER BAR */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '12px', backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF', marginBottom: '24px', alignItems: 'end' }}>
        <div>
          <label style={{ fontSize: '11px', color: '#666', fontWeight: 'bold', display: 'block', marginBottom: '4px' }}>Start date</label>
          <input type="date" value={dateRange.start} onChange={(e) => setDateRange({...dateRange, start: e.target.value})} style={{ width: '90%', padding: '6px', border: '1px solid #CCC', borderRadius: '4px' }} />
        </div>
        <div>
          <label style={{ fontSize: '11px', color: '#666', fontWeight: 'bold', display: 'block', marginBottom: '4px' }}>End date</label>
          <input type="date" value={dateRange.end} onChange={(e) => setDateRange({...dateRange, end: e.target.value})} style={{ width: '90%', padding: '6px', border: '1px solid #CCC', borderRadius: '4px' }} />
        </div>
        <div>
          <label style={{ fontSize: '11px', color: '#666', fontWeight: 'bold', display: 'block', marginBottom: '4px' }}>Brand</label>
          <select value={selectedBrand} onChange={(e) => setSelectedBrand(e.target.value)} style={{ width: '100%', padding: '6px', border: '1px solid #CCC', borderRadius: '4px' }}>
            {uniqueBrands.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: '11px', color: '#666', fontWeight: 'bold', display: 'block', marginBottom: '4px' }}>Region</label>
          <select value={selectedRegion} onChange={(e) => setSelectedRegion(e.target.value)} style={{ width: '100%', padding: '6px', border: '1px solid #CCC', borderRadius: '4px' }}>
            {uniqueRegions.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: '11px', color: '#666', fontWeight: 'bold', display: 'block', marginBottom: '4px' }}>Order type</label>
          <select value={selectedType} onChange={(e) => setSelectedType(e.target.value)} style={{ width: '100%', padding: '6px', border: '1px solid #CCC', borderRadius: '4px' }}>
            {uniqueTypes.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
      </div>

      {/* TAB 1: KPI PERFORMANCE DASHBOARD */}
      {activeTab === 'kpi' && (
        <div>
          {orders.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '60px 0', background: '#FFF', borderRadius: '8px', border: '2px dashed #DDD' }}>
              <p style={{ color: '#666', fontWeight: '600' }}>No dataset imported into AIVI-OPS schema dashboard</p>
              <p style={{ fontSize: '12px', color: '#999' }}>Click "Import Log Sheet" button to populate data layers.</p>
            </div>
          ) : (
            <div>
              {/* TOP KPI CARDS */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '16px', marginBottom: '24px' }}>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <span style={{ fontSize: '12px', color: '#666' }}>Total orders</span>
                  <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#0F766E', marginTop: '6px' }}>{filteredOrders.length}</div>
                </div>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <span style={{ fontSize: '12px', color: '#666' }}>SLA ontime rate</span>
                  <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#DC2626', marginTop: '6px' }}>
                    {Math.round((filteredOrders.filter(o => o.cancelled_qty === 0).length / filteredOrders.length) * 100)}%
                  </div>
                </div>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <span style={{ fontSize: '12px', color: '#666' }}>SLA breached</span>
                  <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#111', marginTop: '6px' }}>
                    {filteredOrders.reduce((a,c) => a + c.cancelled_qty, 0)}
                  </div>
                </div>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <span style={{ fontSize: '12px', color: '#666' }}>Avg processing (OTP)</span>
                  <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#B45309', marginTop: '6px' }}>4h 8m</div>
                </div>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <span style={{ fontSize: '12px', color: '#666' }}>Avg handover (OTH)</span>
                  <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#B45309', marginTop: '6px' }}>10h 7m</div>
                </div>
              </div>

              {/* MAIN CHARTS */}
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr', gap: '20px', marginBottom: '24px' }}>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '16px' }}>Daily orders & SLA rate</h4>
                  <div style={{ height: '260px' }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={getDailyMetrics()}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F0F0F0" />
                        <XAxis dataKey="date" fontSize={10} />
                        <YAxis yAxisId="left" fontSize={10} />
                        <YAxis yAxisId="right" orientation="right" fontSize={10} unit="%" />
                        <Tooltip />
                        <Bar yAxisId="left" dataKey="Orders" fill="#8B4513" barSize={25} />
                        <Line yAxisId="right" type="monotone" dataKey="Ontime %" stroke="#0F766E" strokeWidth={2} />
                      </ComposedChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '16px' }}>Avg processing & handover time (minutes)</h4>
                  <div style={{ height: '260px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999', fontSize: '12px' }}>
                    [OTP / OTH Trend Line View Synchronized]
                  </div>
                </div>
              </div>

              {/* CATEGORICAL BREAKDOWNS */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '16px', marginBottom: '24px' }}>
                <div style={{ backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '13px', fontWeight: 'bold', marginBottom: '12px' }}>By Brand</h4>
                  <ResponsiveContainer width="100%" height="180px">
                    <BarChart data={getCategoricalMetrics('item')}>
                      <XAxis dataKey="name" fontSize={9} />
                      <YAxis fontSize={9} />
                      <Tooltip />
                      <Bar dataKey="Orders" fill="#8B4513" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                
                <div style={{ backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '13px', fontWeight: 'bold', marginBottom: '12px' }}>By Region</h4>
                  <ResponsiveContainer width="100%" height="180px">
                    <BarChart data={getCategoricalMetrics('origin_city')}>
                      <XAxis dataKey="name" fontSize={9} />
                      <YAxis fontSize={9} />
                      <Tooltip />
                      <Bar dataKey="Orders" fill="#8B4513" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={{ backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '13px', fontWeight: 'bold', marginBottom: '12px' }}>By Order Type</h4>
                  <ResponsiveContainer width="100%" height="180px">
                    <BarChart data={getCategoricalMetrics('order_type')}>
                      <XAxis dataKey="name" fontSize={9} />
                      <YAxis fontSize={9} />
                      <Tooltip />
                      <Bar dataKey="Orders" fill="#8B4513" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* TOP AND WORST STORES MATRIX */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '12px', color: '#0F766E' }}>Top 10 Stores by SLA</h4>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid #EEE', textAlign: 'left', color: '#666' }}>
                        <th style={{ padding: '6px 0' }}>STORE</th>
                        <th>ORDERS</th>
                        <th style={{ textAlign: 'right' }}>ONTIME %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getStoreRanking(true).map((s, i) => (
                        <tr key={i} style={{ borderBottom: '1px solid #FAFAFA' }}>
                          <td style={{ padding: '8px 0', fontWeight: '500' }}>{s.name}<br/><span style={{fontSize:'9px', color:'#999'}}>{s.brand} · {s.region}</span></td>
                          <td>{s.Orders}</td>
                          <td style={{ textAlign: 'right', fontWeight: 'bold', color: '#0F766E' }}>{s.rate}%</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                  <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '12px', color: '#DC2626' }}>Worst 10 Stores by SLA</h4>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid #EEE', textAlign: 'left', color: '#666' }}>
                        <th style={{ padding: '6px 0' }}>STORE</th>
                        <th>ORDERS</th>
                        <th>BREACH</th>
                        <th style={{ textAlign: 'right' }}>ONTIME %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getStoreRanking(false).map((s, i) => (
                        <tr key={i} style={{ borderBottom: '1px solid #FAFAFA' }}>
                          <td style={{ padding: '8px 0', fontWeight: '500' }}>{s.name}<br/><span style={{fontSize:'9px', color:'#999'}}>{s.brand} · {s.region}</span></td>
                          <td>{s.Orders}</td>
                          <td>{s.Breached}</td>
                          <td style={{ textAlign: 'right', fontWeight: 'bold', color: '#DC2626' }}>{s.rate}%</td>
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

      {/* TAB 2: DATA VIEW / MASTER RAW DATA TABLE */}
      {activeTab === 'data' && (
        <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h3 style={{ fontSize: '16px', fontWeight: 'bold', margin: 0 }}>Master Operational Database</h3>
            
            {/* SEARCH INPUT */}
            <div style={{ position: 'relative', width: '300px' }}>
              <input 
                type="text" 
                placeholder="Search by Order No, Store, AWB..." 
                value={searchTerm} 
                onChange={(e) => setSearchTerm(e.target.value)} 
                style={{ width: '100%', padding: '8px 12px 8px 32px', border: '1px solid #CCC', borderRadius: '6px', fontSize: '13px' }}
              />
              <Search size={16} color="#999" style={{ position: 'absolute', left: '10px', top: '10px' }} />
            </div>
          </div>

          {/* MASTER DATA TABLE VIEW */}
          <div style={{ overflowX: 'auto', maxHeight: '550px' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
              <thead>
                <tr style={{ background: '#F8FAFC', borderBottom: '2px solid #EAE6DF', position: 'sticky', top: 0, zIndex: 10 }}>
                  {masterColumns.map(col => (
                    <th key={col.id} style={{ padding: '10px 8px', whiteSpace: 'nowrap', color: '#555' }}>{col.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredOrders.length === 0 ? (
                  <tr>
                    <td colSpan={27} style={{ textAlign: 'center', padding: '40px', color: '#999' }}>No matching records found</td>
                  </tr>
                ) : (
                  filteredOrders.map((row, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #FAFAFA', backgroundColor: idx % 2 === 0 ? '#FFF' : '#FDFBF7' }}>
                      <td style={{ padding: '8px', fontWeight: 'bold', color: '#0F766E' }}>{row.order_no}</td>
                      <td style={{ padding: '8px' }}>{row.shipment}</td>
                      <td style={{ padding: '8px' }}>{row.order_id}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.order_date}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.processed_date}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.handover_date}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.inscan_date}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.pickup_date}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.delivered_date}</td>
                      <td style={{ padding: '8px' }}>{row.order_type}</td>
                      <td style={{ padding: '8px' }}>{row.item}</td>
                      <td style={{ padding: '8px' }}>{row.order_qty}</td>
                      <td style={{ padding: '8px' }}>{row.shipped_qty}</td>
                      <td style={{ padding: '8px' }}>{row.rejected_qty}</td>
                      <td style={{ padding: '8px', color: row.cancelled_qty > 0 ? '#DC2626' : '#333' }}>{row.cancelled_qty}</td>
                      <td style={{ padding: '8px' }}>{row.reason}</td>
                      <td style={{ padding: '8px', fontFamily: 'monospace' }}>{row.awb_no}</td>
                      <td style={{ padding: '8px' }}>{row.order_to_process}</td>
                      <td style={{ padding: '8px' }}>{row.order_to_handover}</td>
                      <td style={{ padding: '8px' }}>{row.order_to_inscan}</td>
                      <td style={{ padding: '8px' }}>{row.order_to_pickup}</td>
                      <td style={{ padding: '8px' }}>{row.order_to_delivered}</td>
                      <td style={{ padding: '8px' }}>{row.store_code}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.store_name}</td>
                      <td style={{ padding: '8px', whiteSpace: 'nowrap' }}>{row.mall_name}</td>
                      <td style={{ padding: '8px' }}>{row.origin_city}</td>
                      <td style={{ padding: '8px' }}>{row.origin_country}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

    </div>
  );
}
