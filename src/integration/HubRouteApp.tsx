import { routeDocuments } from './routeDocuments';
import printAssets from 'virtual:routeready-print-assets';
import { GenerationArchiver } from './generatedArtifact';
import type { PullSheetDocument } from '../types/pullSheet';
import { useEffect, useMemo, useState, useCallback } from 'react';
import App from '../App';
import { ArchivePanel } from './ArchivePanel';
import { ArchiveClient, facilityBase, type RouteContext } from './archiveClient';
import './hub-route.css';
export function HubRouteApp(){
 const [operations,setOperations]=useState({preparationDate:'',shippingDate:'',trailer:''});
 const [context,setContext]=useState<RouteContext|null>(null),[error,setError]=useState('');
 const client=useMemo(()=>new ArchiveClient(facilityBase(window.location.pathname)),[]);
 const archiver=useMemo(()=>new GenerationArchiver(artifact=>client.saveArtifact(artifact),printAssets),[client]);
 const onGenerated=useCallback((document:PullSheetDocument,markup:string)=>{
  try{const scoped=routeDocuments(document),metadata={preparationDate:operations.preparationDate||null,shippingDate:operations.shippingDate||null,trailer:operations.trailer.trim()||null};
   archiver.generated(document,markup,metadata);for(const route of scoped)archiver.generated(route.document,route.markup,metadata);
  }catch(error){setError((error as Error).message);}
 },[archiver,operations]);
 useEffect(()=>{let active=true;async function verify(){try{const value=await client.session();if(active){setContext(value);setError('');}}catch(error){if(active){setContext(null);setError((error as Error).message);}}}
 void verify();const focus=()=>void verify();const timer=window.setInterval(()=>{if(document.visibilityState==='visible')void verify();},30000);window.addEventListener('focus',focus);return()=>{active=false;window.clearInterval(timer);window.removeEventListener('focus',focus);};},[client]);
 if(!context)return <main className="hub-route"><h1>SCRouteReady</h1><p role={error?'alert':'status'}>{error||'Verifying your Hub session and active facility...'}</p><a href="/">Return to OpsReady Hub</a></main>;
 return <div className="hub-route" key={context.facility.id}><div className="route-archive-ui hub-route-context"><a href="/">OpsReady Hub</a><span>{context.facility.name}</span><span>{context.creator.name}</span></div>
 <section className="route-archive-ui hub-route-context" aria-label="Route operational metadata"><strong>Operational dates for newly generated archives</strong><label>Preparation / warehouse work date<input type="date" value={operations.preparationDate} onChange={e=>setOperations(o=>({...o,preparationDate:e.target.value}))}/></label><label>Ship / route date<input type="date" value={operations.shippingDate} onChange={e=>setOperations(o=>({...o,shippingDate:e.target.value}))}/></label><label>Trailer (only when common to this generated document)<input value={operations.trailer} maxLength={100} onChange={e=>setOperations(o=>({...o,trailer:e.target.value}))}/></label><small>Set metadata before generating. Each generation snapshots these values; changing them later does not alter an existing archive. Blank values remain unknown. Existing archived output is unchanged.</small></section>
 {error&&<p role="alert">{error}</p>}
 <App onGenerated={onGenerated} renderArchive={()=><ArchivePanel key={context.facility.id} client={client} context={context} archiver={archiver}/>}/></div>;
}
