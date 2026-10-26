import type { PullSheetDocument } from '../types/pullSheet';
import type { ArchiveMetadata } from './archiveClient';
export type GeneratedArtifact=Readonly<{id:string;html:string;metadata:ArchiveMetadata}>;
export function captureGeneratedArtifact(document:PullSheetDocument,markup:string,assets:{css:string;logo:string},id:string=crypto.randomUUID(),operations?:Omit<ArchiveMetadata,'routeIds'>):GeneratedArtifact{
 if(!markup.startsWith('<div class="print-document"')||!assets.logo.startsWith('data:image/png;base64,'))throw new Error('Complete print output was not available.');
 const body=markup.split('/assets/SCRouteReady_logo_cropped.png').join(assets.logo);
 const html='<!doctype html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="routeready-print-artifact" content="v1"><title>Archived Pull Sheet</title><style data-print-css="v1">'+assets.css+'</style></head><body><main class="app-shell"><div class="workspace-grid"><div class="primary-column"><section class="preview-shell">'+body+'</section></div></div></main></body></html>';
 // Optional integrated operations metadata is explicit; never infer dates/trailer from route text.
 const metadata:ArchiveMetadata={routeIds:[...new Set(document.routes.map(route=>route.routeDisplay))],trailer:operations?.trailer??null,preparationDate:operations?.preparationDate??null,shippingDate:operations?.shippingDate??null};
 Object.freeze(metadata.routeIds);Object.freeze(metadata);
 return Object.freeze({id,html,metadata});
}
export type ArchiveJob={id:string;artifact:GeneratedArtifact|null;routeIds:string[];status:'Saving'|'Saved'|'Not saved';error:string;markup:string;document:PullSheetDocument|null;operations?:Omit<ArchiveMetadata,'routeIds'>};
export class GenerationArchiver{
 private byDocument=new WeakMap<PullSheetDocument,ArchiveJob>();
 private items:ArchiveJob[]=[];private listeners=new Set<()=>void>();private flights=new Map<string,Promise<void>>();
 constructor(private save:(artifact:GeneratedArtifact)=>Promise<unknown>,private assets:{css:string;logo:string}){this.assets={...assets};}
 jobs(){return [...this.items];}
 subscribe(listener:()=>void){this.listeners.add(listener);return()=>{this.listeners.delete(listener);};}
 private emit(){for(const listener of this.listeners)listener();}
 generated(document:PullSheetDocument,markup:string,operations?:Omit<ArchiveMetadata,'routeIds'>){
  const existing=this.byDocument.get(document);if(existing)return existing;
  const job:ArchiveJob={id:crypto.randomUUID(),artifact:null,routeIds:document.routes.map(r=>r.routeDisplay),status:'Saving',error:'',markup,document,operations:operations?{...operations}:undefined};
  this.byDocument.set(document,job);this.items.unshift(job);this.emit();
  // Capture strings from the completed DOM; storage never joins the generation
  // promise or changes print availability. Retain failures for one-click retry.
  void this.retry(job);return job;
 }
 retry(job:ArchiveJob):Promise<void>{
  if(job.status==='Saved')return Promise.resolve();const running=this.flights.get(job.id);if(running)return running;
  job.status='Saving';job.error='';this.emit();
  const work=Promise.resolve().then(async()=>{
   try{job.artifact??=captureGeneratedArtifact(job.document!,job.markup,this.assets,job.id,job.operations);job.document=null;job.markup='';await this.save(job.artifact);job.artifact=null;job.status='Saved';}
   catch(error){job.status='Not saved';job.error=(error as Error).message;}
   finally{this.flights.delete(job.id);this.emit();}
  });this.flights.set(job.id,work);return work;
 }
}
