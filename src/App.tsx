import VAStaleAccess from './va_stale_access';

export default function App() {
  return (
    <div style={{minHeight:"100vh",background:"#f4f6f9",fontFamily:"system-ui,-apple-system,sans-serif"}}>
      {/* Header */}
      <header style={{background:"#112240",color:"white",padding:"0 2rem",display:"flex",alignItems:"center",gap:16,height:56,boxShadow:"0 2px 8px rgba(0,0,0,0.18)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:32,height:32,borderRadius:6,background:"#C8102E",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:14,letterSpacing:-0.5}}>VA</div>
          <div>
            <div style={{fontSize:15,fontWeight:600,lineHeight:1.2}}>VHA WMC Access Review</div>
            <div style={{fontSize:11,color:"#8aa4c8",lineHeight:1.2}}>HR-Smart System Access Compliance</div>
          </div>
        </div>
        <div style={{flex:1}}/>
        <div style={{fontSize:12,color:"#8aa4c8"}}>System: HR-Smart · Role: VA Admin Officer</div>
      </header>

      {/* Main */}
      <main style={{maxWidth:1280,margin:"0 auto",padding:"1.5rem 2rem"}}>
        <VAStaleAccess/>
      </main>
    </div>
  );
}
