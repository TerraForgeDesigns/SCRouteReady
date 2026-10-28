import { renderToStaticMarkup } from 'react-dom/server';
import type { PullSheetDocument } from '../types/pullSheet';
import { PullSheetPreview } from '../components/PullSheetPreview';
const cache=new WeakMap<PullSheetDocument,{document:PullSheetDocument;markup:string}[]>();
/** Render existing print components with only this route and its own pallet keys. */
export function routeDocuments(source:PullSheetDocument){
 const existing=cache.get(source);if(existing)return existing;
 const result=source.routes.length<2?[]:source.routes.map(route=>{
  const document={...source,routes:[route],palletKeyPages:source.palletKeyPages.filter(page=>page.routeKey===route.routeKey)};
  const preview=renderToStaticMarkup(<PullSheetPreview document={document}/>);
  const start=preview.indexOf('<div class="print-document"');
  if(start<0||!preview.endsWith('</section>'))throw new Error('Route-specific print output could not be prepared. Generate again before assigning this batch.');
  return {document,markup:preview.slice(start,-'</section>'.length)};
 });
 cache.set(source,result);return result;
}
