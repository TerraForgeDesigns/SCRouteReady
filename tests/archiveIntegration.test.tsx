import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { renderToStaticMarkup } from 'react-dom/server';
import * as XLSX from 'xlsx';
import App from '../src/App';
import { PullSheetPreview } from '../src/components/PullSheetPreview';
import { generatePullSheetDocument } from '../src/services/pullSheetGenerator';
import { ArchivePanel } from '../src/integration/ArchivePanel';
import { ArchiveClient,facilityBase } from '../src/integration/archiveClient';
import { GenerationArchiver,captureGeneratedArtifact,type GeneratedArtifact } from '../src/integration/generatedArtifact';
const assets={css:['app.css','pull-sheet.css','print.css'].map(name=>readFileSync('src/styles/'+name,'utf8')).join('\n'),logo:'data:image/png;base64,'+readFileSync('public/assets/SCRouteReady_logo_cropped.png','base64')};
function generated(){
 const rows=[{ '_internalrowindex':1,'Route from Descartes':'Route 12','Customer Type':'Trade','Stop #':2,'Pick Ticket No':101,'Po No':'PO-1',Qty:2,UOM:'EA',Bin:'A1','DMAS Part#':'PART-2','P21 Item Id':'ITEM','Item Desc':'Printed item','Ship To Name':'Customer A','Ship To City':'City','Ship To State':'SC','Order No':'ORDER','Route Code':'R12'},
 {'_internalrowindex':2,'Route from Descartes':'Route 12','Customer Type':'Trade','Stop #':2,'Pick Ticket No':102,'Po No':'PO-2',Qty:1,UOM:'EA',Bin:'A2','DMAS Part#':'PART-3','P21 Item Id':'ITEM2','Item Desc':'Other printed item','Ship To Name':'Customer B','Ship To City':'Town','Ship To State':'SC','Order No':'ORDER2','Route Code':'R12'}];
 const workbook=XLSX.utils.book_new();XLSX.utils.book_append_sheet(workbook,XLSX.utils.json_to_sheet(rows),'Results');
 const document=generatePullSheetDocument(workbook),preview=renderToStaticMarkup(<PullSheetPreview document={document}/>);
 const markup=preview.slice(preview.indexOf('<div class="print-document"'),-'</section>'.length);return {document,markup};
}
test('standalone keeps original Print/generation actions with no automatic Hub calls or manual archive form',()=>{
 const html=renderToStaticMarkup(<App/>);for(const text of ['Generate Pull Sheet','Print','Finish &amp; Clear','Nothing is transmitted or intentionally stored.'])assert.equal(html.includes(text),true);assert.equal(html.includes('Recent Pull Sheets'),false);
});
test('captures the complete existing route/key print markup with exact CSS and embedded logo, independent of future updates',()=>{
 const {document,markup}=generated(),artifact=captureGeneratedArtifact(document,markup,assets);
 assert.equal(document.palletKeyPages.length,1);for(const text of ['Customer A','Customer B','PART-2','PART-3','pallet-key-page','designation-letter','with-key-column'])assert.equal(artifact.html.includes(text),true);
 assert.equal(artifact.html.includes(assets.css),true);assert.equal(artifact.html.includes(assets.logo),true);assert.equal(artifact.html.includes('size: letter landscape'),true);assert.equal(artifact.html.includes('/assets/'),false);
 const original=artifact.html;document.routes[0].rows[0].itemDescription='Future change';const laterAssets={css:'future CSS',logo:assets.logo};captureGeneratedArtifact(document,markup,laterAssets);
 assert.equal(artifact.html,original);assert.equal(artifact.html.includes('Future change'),false);assert.deepEqual(artifact.metadata,{routeIds:['Route 12'],trailer:null,preparationDate:null,shippingDate:null});
});
test('queues automatic saving once per successful generation and never waits for storage to make printing available',async()=>{
 const {document,markup}=generated();let calls=0,finish:()=>void=()=>{};
 const archiver=new GenerationArchiver(async()=>{calls++;await new Promise<void>(resolve=>{finish=resolve;});},assets);
 const first=archiver.generated(document,markup);assert.equal(first.status,'Saving');assert.equal(calls,0);
 assert.equal(archiver.generated(document,markup),first);await Promise.resolve();assert.equal(calls,1);const running=archiver.retry(first);finish();await running;assert.equal(first.status,'Saved');await archiver.retry(first);assert.equal(calls,1);
 const next=generated();const second=archiver.generated(next.document,next.markup);assert.equal(first.id===second.id,false);await Promise.resolve();finish();await archiver.retry(second);assert.equal(calls,2);
});
test('a failed save shows Not saved with one-click retry of the same immutable content and generation ID',async()=>{
 const {document,markup}=generated(),sent:GeneratedArtifact[]=[];let fail=true;
 const archiver=new GenerationArchiver(async artifact=>{sent.push(artifact);if(fail)throw new Error('Storage unavailable');},assets);
 const job=archiver.generated(document,markup);await archiver.retry(job);assert.equal(job.status,'Not saved');
 const html=renderToStaticMarkup(<ArchivePanel client={new ArchiveClient('/routes/facilities/f/api')} context={{facility:{id:'f',name:'Authorized DC'},creator:{id:'e',name:'Verified'},role:'DC_COORDINATOR',csrfToken:'csrf',archiveAvailable:false}} archiver={archiver}/>);
 for(const text of ['Not saved','Retry save','Printing is still available','Recent Pull Sheets','Saved pull sheet dates'])assert.equal(html.includes(text),true);
 for(const removed of ['type="file"','Finalize and save','Trailer (when available)','Shipping date (when available)'])assert.equal(html.includes(removed),false);
 document.routes[0].rows[0].itemDescription='Changed later';fail=false;await archiver.retry(job);assert.equal(job.status,'Saved');assert.equal(sent[0],sent[1]);assert.equal(sent[0].id,sent[1].id);assert.equal(sent[0].html.includes('Changed later'),false);
});
test('the archive client uses stable artifact IDs, authenticated date-first lists and rejects unconfirmed saves',async()=>{
 const calls:{url:string;init:RequestInit}[]=[];let confirmed=false;
 const client=new ArchiveClient('/routes/facilities/f/api',async(input,init={})=>{calls.push({url:String(input),init});return Response.json(String(input).endsWith('/session')?{csrfToken:'csrf'}:init.method?{saved:confirmed}:{dates:[],pullSheets:[],next:null});});
 await client.session();const {document,markup}=generated(),artifact=captureGeneratedArtifact(document,markup,assets);let rejected=false;try{await client.saveArtifact(artifact);}catch{rejected=true;}assert.equal(rejected,true);confirmed=true;await client.saveArtifact(artifact);
 assert.equal(calls[1].url,calls[2].url);assert.equal(calls[1].init.body,calls[2].init.body);assert.equal(calls[2].init.credentials,'same-origin');assert.equal((calls[2].init.headers as Record<string,string>)['X-CSRF-Token'],'csrf');
 await client.dates('recent');await client.documents('recent','2026-10-21');assert.equal(calls[3].url.includes('/pull-sheet-dates?'),true);assert.equal(calls[4].url.includes('date=2026-10-21'),true);
 assert.equal(facilityBase('/routes/facilities/f/'),'/routes/facilities/f/api');assert.throws(()=>facilityBase('/?facilityId=f'));
});

test('saves explicit integrated operational metadata without changing generated routes or print markup',async()=>{
 const {document,markup}=generated(),operations={preparationDate:'2026-10-25',shippingDate:'2026-10-26',trailer:'DEV-TRAILER'};
 const captured:GeneratedArtifact[]=[];const archiver=new GenerationArchiver(async artifact=>{captured.push(artifact);return {saved:true};},assets);
 const job=archiver.generated(document,markup,operations);operations.shippingDate='2026-12-31';await archiver.retry(job);
 assert.equal(job.status,'Saved');assert.equal(captured[0].metadata.preparationDate,'2026-10-25');assert.equal(captured[0].metadata.shippingDate,'2026-10-26');assert.equal(captured[0].metadata.trailer,'DEV-TRAILER');
 const plain=captureGeneratedArtifact(document,markup,assets);assert.equal(plain.html,captured[0].html);assert.equal(plain.metadata.shippingDate,null);assert.deepEqual(plain.metadata.routeIds,captured[0].metadata.routeIds);
});
