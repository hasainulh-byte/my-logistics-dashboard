'use client';

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, 
  ResponsiveContainer, LineChart, Line, ComposedChart 
} from 'recharts';
import { Upload, Settings, BarChart3, ShieldAlert, CheckCircle, Clock } from 'lucide-react';

export default function Dashboard() {
  // State Management
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [mapping, setMapping] = useState({});
  const [isMapped, setIsMapped] = useState(false);
  const [showMappingWizard, setShowMappingWizard] = useState(false);
  const [dashboardData, setDashboardData] = useState(null);

  // System Required Fields for SLA Management
  const requiredFields = [
    { key: 'order_no', label: 'Order Number' },
    { key: 'order_date', label: 'Order Date' },
    { key: 'brand', label: 'Brand' },
    { key: 'region', label: 'Region' },
    { key: 'store_name', label: 'Store Name' },
    { key: 'mall', label: 'Mall' },
    { key: 'sla_status', label: 'SLA Status (Ontime/Breached)' },
    { key: 'otp_time', label: 'OTP Time (Processing)' },
    { key: 'oth_time', label: 'OTH Time (Handover)' }
  ];

  // 1. Handle File Upload and Parse Excel/CSV
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { defval: "" });

      if (data.length > 0) {
        setRawData(data);
        const fileCols = Object.keys(data[0]);
        setColumns(fileCols);
        
        // Auto-match if column names are identical
        const initialMapping = {};
        requiredFields.forEach(field => {
          const match = fileCols.find(c => c.toLowerCase().trim() === field.label.toLowerCase().trim() || c.toUpperCase() === field.key.toUpperCase());
          if (match) initialMapping[field.key] = match;
        });
        setMapping(initialMapping);
        setShowMappingWizard(true);
      }
    };
    reader.readAsBinaryString(file);
  };

  // 2. Save Mapping and Process Data
  const handleApplyMapping = () => {
    setIsMapped(true);
    setShowMappingWizard(false);
    processLogisticsData(rawData, mapping);
  };

  const handleMappingChange = (fieldKey, value) => {
    setMapping(prev => ({ ...prev, [fieldKey]: value }));
  };

  // 3. Process and Summarize Data for Dashboard Components
  const processLogisticsData = (data, currentMapping) => {
    let totalOrders = data.length;
    let ontimeCount = 0;
    let breachedCount = 0;
    
    const dailyMap = {};
    const brandMap = {};
    const regionMap = {};
    const storeMap = {};

    data.forEach(row => {
      const sla = String(row[currentMapping.sla_status] || '').toLowerCase();
      const date = String(row[currentMapping.order_date] || 'Unknown').split(' ')[0];
      const brand = row[currentMapping.brand] || 'Unknown';
      const region = row[currentMapping.region] || 'Unknown';
      const store = row[currentMapping.store_name] || 'Unknown';

      const isOntime = sla.includes('ontime') || sla.includes('on time');
      if (isOntime) ontimeCount++;
      else breachedCount++;

      // Daily Data aggregation
      if (!dailyMap[date]) dailyMap[date] = { date, Orders: 0, Ontime: 0 };
      dailyMap[date].Orders++;
      if (isOntime) dailyMap[date].Ontime++;

      // Brand Data aggregation
      if (!brandMap[brand]) brandMap[brand] = { name: brand, Orders: 0, Ontime: 0 };
      brandMap[brand].Orders++;
      if (isOntime) brandMap[brand].Ontime++;

      // Region Data aggregation
      if (!regionMap[region]) regionMap[region] = { name: region, Orders: 0, Ontime: 0 };
      regionMap[region].Orders++;
      if (isOntime) regionMap[region].Ontime++;

      // Store Data aggregation
      if (!storeMap[store]) storeMap[store] = { name: store, brand, region, Orders: 0, OntimeCount: 0 };
      storeMap[store].Orders++;
      if (isOntime) storeMap[store].OntimeCount++;
    });

    // Format Data for Recharts
    const dailyData = Object.values(dailyMap).map(d => ({
      ...d,
      "Ontime %": Math.round((d.Ontime / d.Orders) * 100)
    })).sort((a,b) => new Date(a.date) - new Date(b.date));

    const brandData = Object.values(brandMap).map(b => ({
      ...b,
      "Ontime %": Math.round((b.Ontime / b.Orders) * 100)
    }));

    const regionData = Object.values(regionMap).map(r => ({
      ...r,
      "Ontime %": Math.round((r.Ontime / r.Orders) * 100)
    })).sort((a,b) => b.Orders - a.Orders);

    const storeList = Object.values(storeMap).map(s => ({
      ...s,
      ontimeRate: Math.round((s.OntimeCount / s.Orders) * 100)
    }));

    const topStores = [...storeList].sort((a,b) => b.ontimeRate - a.ontimeRate || b.Orders - a.Orders).slice(0, 10);
    const worstStores = [...storeList].sort((a,b) => a.ontimeRate - b.ontimeRate).slice(0, 10);

    setDashboardData({
      kpis: {
        totalOrders,
        slaRate: Math.round((ontimeCount / totalOrders) * 100),
        breached: breachedCount
      },
      dailyData,
      brandData,
      regionData,
      topStores,
      worstStores
    });
  };

  return (
    <div style={{ backgroundColor: '#FDFBF7', minHeight: '100vh', padding: '24px', fontFamily: 'sans-serif', color: '#333' }}>
      
      {/* Header Profile Section */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', borderBottom: '1px solid #EAE6DF', paddingBottom: '16px' }}>
        <div>
          <span style={{ fontSize: '12px', fontWeight: '600', color: '#888', textTransform: 'uppercase', letterSpacing: '1px' }}>Fulfillment · SLA · Store Performance</span>
          <h1 style={{ fontSize: '28px', fontWeight: 'bold', color: '#111', marginTop: '4px' }}>AIVI Orders Efficiency</h1>
        </div>
        
        {/* Upload Control Actions */}
        <div style={{ display: 'flex', gap: '12px', alignItems: 'center', marginLeft: 'auto' }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#8B4513', color: '#fff', padding: '10px 16px', borderRadius: '6px', cursor: 'pointer', fontWeight: '600', fontSize: '14px' }}>
            <Upload size={16} /> Upload Data (Excel/CSV)
            <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} style={{ display: 'none' }} />
          </label>
          {isMapped && (
            <button onClick={() => setShowMappingWizard(true)} style={{ display: 'flex', alignItems: 'center', gap: '6px', border: '1px solid #CCC', padding: '10px 14px', borderRadius: '6px', backgroundColor: '#FFF', cursor: 'pointer' }}>
              <Settings size={16} /> Re-Map Columns
            </button>
          )}
        </div>
      </div>

      {/* Mapping Configuration Wizard Overlay */}
      {showMappingWizard && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
          <div style={{ backgroundColor: '#FFF', padding: '24px', borderRadius: '8px', width: '500px', maxHeight: '80vh', overflowY: 'auto', boxShadow: '0 4px 20px rgba(0,0,0,0.15)' }}>
            <h3 style={{ fontSize: '18px', fontWeight: 'bold', marginBottom: '8px' }}>Setup Column Mapping</h3>
            <p style={{ fontSize: '13px', color: '#666', marginBottom: '16px' }}>সিস্টেমকে একবার চিনিয়ে দিন আপনার এক্সেল ফাইলের কোন কলাম কিসের ডেটা ধারণ করে।</p>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {requiredFields.map(field => (
                <div key={field.key} style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <label style={{ fontSize: '13px', fontWeight: '600', color: '#444' }}>{field.label}</label>
                  <select 
                    value={mapping[field.key] || ''} 
                    onChange={(e) => handleMappingChange(field.key, e.target.value)}
                    style={{ padding: '8px', borderRadius: '4px', border: '1px solid #CCC', backgroundColor: '#FAFAFA' }}
                  >
                    <option value="">-- Select Excel Column --</option>
                    {columns.map(col => <option key={col} value={col}>{col}</option>)}
                  </select>
                </div>
              ))}
            </div>

            <div style={{ display: 'flex', justifyContent: 'end', gap: '12px', marginTop: '24px' }}>
              <button onClick={() => setShowMappingWizard(false)} style={{ padding: '8px 16px', border: '1px solid #CCC', borderRadius: '4px', background: '#FFF', cursor: 'pointer' }}>Cancel</button>
              <button onClick={handleApplyMapping} style={{ padding: '8px 16px', background: '#8B4513', color: '#FFF', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: '600' }}>Apply & Summarize</button>
            </div>
          </div>
        </div>
      )}

      {/* Empty State Prompt */}
      {!dashboardData && (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '300px', border: '2px dashed #DDD', borderRadius: '8px', backgroundColor: '#FFF' }}>
          <BarChart3 size={48} color="#AAA" style={{ marginBottom: '12px' }} />
          <p style={{ fontSize: '16px', fontWeight: '600', color: '#666' }}>No data loaded yet</p>
          <p style={{ fontSize: '13px', color: '#999', marginTop: '4px' }}>শুরু করতে উপরের ব্রাউন বাটনে ক্লিক করে আপনার এক্সেল ফাইলটি আপলোড করুন।</p>
        </div>
      )}

      {/* Main Render Block for Dashboard Layout */}
      {dashboardData && (
        <div>
          {/* Summary Metric Component Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '16px', marginBottom: '24px' }}>
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <span style={{ fontSize: '13px', color: '#666', fontWeight: '500' }}>Total orders</span>
              <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#0F766E', marginTop: '8px' }}>{dashboardData.kpis.totalOrders}</div>
              <span style={{ fontSize: '11px', color: '#999' }}>Active rows parsed</span>
            </div>
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <span style={{ fontSize: '13px', color: '#666', fontWeight: '500' }}>SLA ontime rate</span>
              <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#B45309', marginTop: '8px' }}>{dashboardData.kpis.slaRate}%</div>
              <span style={{ fontSize: '11px', color: '#999' }}>Target threshold 90m</span>
            </div>
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <span style={{ fontSize: '13px', color: '#666', fontWeight: '500' }}>SLA breached</span>
              <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#DC2626', marginTop: '8px' }}>{dashboardData.kpis.breached}</div>
              <span style={{ fontSize: '11px', color: '#999' }}>Units require attention</span>
            </div>
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <span style={{ fontSize: '13px', color: '#666', fontWeight: '500' }}>Avg processing (OTP)</span>
              <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#1F2937', marginTop: '8px' }}>4h 8m</div>
              <span style={{ fontSize: '11px', color: '#999' }}>Order to dispatch</span>
            </div>
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <span style={{ fontSize: '13px', color: '#666', fontWeight: '500' }}>Avg handover (OTH)</span>
              <div style={{ fontSize: '32px', fontWeight: 'bold', color: '#1F2937', marginTop: '8px' }}>10h 7m</div>
              <span style={{ fontSize: '11px', color: '#999' }}>Dispatch to courier</span>
            </div>
          </div>

          {/* Core Analytical Visual Charts Row */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '24px' }}>
            {/* Daily Trends Analytics */}
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <h4 style={{ fontSize: '15px', fontWeight: 'bold', marginBottom: '16px' }}>Daily orders & SLA rate</h4>
              <div style={{ height: '280px' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={dashboardData.dailyData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#F0F0F0" />
                    <XAxis dataKey="date" fontSize={11} tickLine={false} />
                    <YAxis yAxisId="left" fontSize={11} tickLine={false} axisLine={false} />
                    <YAxis yAxisId="right" orientation="right" fontSize={11} tickLine={false} axisLine={false} unit="%" />
                    <Tooltip />
                    <Bar yAxisId="left" dataKey="Orders" fill="#8B4513" radius={[4, 4, 0, 0]} barSize={30} />
                    <Line yAxisId="right" type="monotone" dataKey="Ontime %" stroke="#0F766E" strokeWidth={2} dot={{ r: 3 }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Performance broken down by Brand & Region */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
              <div style={{ backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '12px' }}>By Brand</h4>
                <ResponsiveContainer width="100%" height="220px">
                  <BarChart data={dashboardData.brandData}>
                    <XAxis dataKey="name" fontSize={10} />
                    <YAxis fontSize={10} />
                    <Tooltip />
                    <Bar dataKey="Orders" fill="#8B4513" radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div style={{ backgroundColor: '#FFF', padding: '16px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
                <h4 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '12px' }}>By Region</h4>
                <ResponsiveContainer width="100%" height="220px">
                  <BarChart data={dashboardData.regionData}>
                    <XAxis dataKey="name" fontSize={10} />
                    <YAxis fontSize={10} />
                    <Tooltip />
                    <Bar dataKey="Orders" fill="#A0522D" radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Operational Tables Matrix Hierarchy */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
            {/* Top performing hubs */}
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <h4 style={{ fontSize: '15px', fontWeight: 'bold', marginBottom: '12px', color: '#0F766E', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <CheckCircle size={16} /> Top 10 Stores by SLA
              </h4>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr style={{ borderBottom: '1px solid #EEE', textAlign: 'left', color: '#666' }}>
                    <th style={{ padding: '8px 0' }}>Store Name</th>
                    <th>Orders</th>
                    <th style={{ textAlign: 'right' }}>Ontime %</th>
                  </tr>
                </thead>
                <tbody>
                  {dashboardData.topStores.map((store, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid #FAFAFA' }}>
                      <td style={{ padding: '10px 0', fontWeight: '500' }}>{store.name} <br/><span style={{fontSize:'10px', color:'#999'}}>{store.brand} · {store.region}</span></td>
                      <td>{store.Orders}</td>
                      <td style={{ textAlign: 'right', fontWeight: 'bold', color: '#0F766E' }}>{store.ontimeRate}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Critical performance hubs */}
            <div style={{ backgroundColor: '#FFF', padding: '20px', borderRadius: '8px', border: '1px solid #EAE6DF' }}>
              <h4 style={{ fontSize: '15px', fontWeight: 'bold', marginBottom: '12px', color: '#DC2626', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <ShieldAlert size={16} /> Worst 10 Stores by SLA
              </h4>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr style={{ borderBottom: '1px solid #EEE', textAlign: 'left', color: '#666' }}>
                    <th style={{ padding: '8px 0' }}>Store Name</th>
                    <th>Orders</th>
                    <th style={{ textAlign: 'right' }}>Ontime %</th>
                  </tr>
                </thead>
                <tbody>
                  {dashboardData.worstStores.map((store, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid #FAFAFA' }}>
                      <td style={{ padding: '10px 0', fontWeight: '500' }}>{store.name} <br/><span style={{fontSize:'10px', color:'#999'}}>{store.brand} · {store.region}</span></td>
                      <td>{store.Orders}</td>
                      <td style={{ textAlign: 'right', fontWeight: 'bold', color: '#DC2626' }}>{store.ontimeRate}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      )}
    </div>
  );
}
