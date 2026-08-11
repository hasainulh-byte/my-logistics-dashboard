export const metadata = {
  title: 'AIVI Orders Efficiency Dashboard',
  description: 'Logistics SLA and Performance Tracking Dashboard',
}

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0, padding: 0, backgroundColor: '#FDFBF7' }}>
        {children}
      </body>
    </html>
  )
}
