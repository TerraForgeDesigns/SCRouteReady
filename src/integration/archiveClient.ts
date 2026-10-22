import type { GeneratedArtifact } from './generatedArtifact';
export type ArchiveMetadata={routeIds:string[];trailer:string|null;preparationDate:string|null;shippingDate:string|null};
export type PullSheetArchive=ArchiveMetadata&{id:string;facilityId:string;creator:{id:string;name:string};createdAt:string;byteLength:number;saved:boolean;format?:"pdf"|"print-html-v1"};
export type RouteContext={facility:{id:string;name:string};creator:{id:string;name:string};role:string;csrfToken:string;archiveAvailable:boolean};
export type ArchivePage={pullSheets:PullSheetArchive[];cutoff:string;next:string|null};
export class ArchiveError extends Error {constructor(public status:number,message:string){super(message);}}
export function facilityBase(pathname:string):string {
 const match=/^\/routes\/facilities\/([a-zA-Z0-9_-]{1,64})(?:\/|$)/.exec(pathname);
 if(!match)throw new Error('Open RouteReady from your active facility in the Hub.');
 return `/routes/facilities/${match[1]}/api`;
}
export class ArchiveClient {
 private csrf='';
 constructor(private base:string,private transport:typeof fetch=(...args)=>fetch(...args)){}
 private async request<T>(path:string,init:RequestInit={}):Promise<T>{
  const response=await this.transport(this.base+path,{...init,credentials:'same-origin',cache:'no-store'});
  const data=await response.json();
  if(!response.ok)throw new ArchiveError(response.status,data?.error?.message??'Archive request failed. Printing is still available.');
  return data as T;
 }
 async session(){const context=await this.request<RouteContext>('/session');this.csrf=context.csrfToken;return context;}
 list(period:'recent'|'archive',q='',before:string|null=null){const params=new URLSearchParams({period,q});if(before)params.set('before',before);return this.request<ArchivePage>('/pull-sheets?'+params);}
 async saveArtifact(artifact:GeneratedArtifact){
  if(!this.csrf)throw new ArchiveError(401,'Reopen RouteReady from the Hub before retrying.');
  const result=await this.request<PullSheetArchive>(`/pull-sheets/${encodeURIComponent(artifact.id)}/artifact`,{method:'PUT',headers:{'Content-Type':'text/html','X-CSRF-Token':this.csrf,'X-Pull-Sheet-Metadata':encodeURIComponent(JSON.stringify(artifact.metadata))},body:artifact.html});
  if(result.saved!==true)throw new Error('Not saved. The server did not confirm archival storage.');return result;
 }
 artifactUrl(id:string){return `${this.base}/pull-sheets/${encodeURIComponent(id)}/artifact`;}
 dates(period:'recent'|'archive',q='',before:string|null=null){const params=new URLSearchParams({period,q});if(before)params.set('beforeDate',before);return this.request<{dates:{date:string;count:number}[];next:string|null}>('/pull-sheet-dates?'+params);}
 documents(period:'recent'|'archive',date:string,q='',before:string|null=null){const params=new URLSearchParams({period,date,q});if(before)params.set('before',before);return this.request<ArchivePage>('/pull-sheets?'+params);}

 pdfUrl(id:string){return `${this.base}/pull-sheets/${encodeURIComponent(id)}/pdf`;}
}
