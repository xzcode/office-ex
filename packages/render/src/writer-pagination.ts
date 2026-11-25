// SPDX-License-Identifier: MIT
import {assert} from '@office-ex/contracts';
import {freezeData,paragraphBoundaries,type DocumentSnapshot,type DeepReadonly,type WriterDocument,type WriterSection} from '@office-ex/core';
import type {FontMetrics} from './index.js';
import {layoutWriter,type WriterLayout,type WriterLayoutOptions,type WriterPage} from './writer-layout.js';

/** A section is an independent page sequence. Immutable section identity and
 * style/resource identity permit reuse; changed sections reflow and retain
 * identical pages. The paragraph shaper separately caches native line boxes.
 * Layout budgets apply globally even on cache hits. */
interface Entry {
  source:DeepReadonly<WriterSection>;
  defaults:DeepReadonly<WriterDocument['defaultTextStyle']>;
  styles:DeepReadonly<WriterDocument['styles']>;
  resources:DeepReadonly<WriterDocument['resources']>;
  layout:DeepReadonly<WriterLayout>;
  clusters:number;
  blocks:number;
}
export interface PaginationStats {sectionsComputed:number;sectionsReused:number;pagesReused:number}
function cost(section:DeepReadonly<WriterSection>):{clusters:number;blocks:number}{
  let clusters=0,blocks=0;
  for(const b of section.blocks){
    if(b.type==='paragraph'){clusters+=paragraphBoundaries(b).length-1;blocks++;}
    else if(b.type==='table')for(const row of b.rows){blocks++;for(const cell of row.cells)for(const p of cell.blocks){assert(p.type==='paragraph','unsupported','Nested table layout not supported');blocks++;clusters+=paragraphBoundaries(p).length-1;}}
    else blocks++;
  }
  return {clusters,blocks};
}
export class WriterPaginationEngine {
  private entries=new Map<string,Entry>();
  private signature='';
  stats:PaginationStats={sectionsComputed:0,sectionsReused:0,pagesReused:0};
  clear():void{this.entries.clear();this.signature='';}
  layout(snapshot:DocumentSnapshot,provider:FontMetrics,options:WriterLayoutOptions={}):DeepReadonly<WriterLayout>{
    assert(snapshot.document.kind==='writer','invalid','Expected Writer document');
    const model=snapshot.document,maxPages=options.maxPages??100,maxClusters=options.maxClusters??30000;
    assert(Number.isSafeInteger(maxPages)&&maxPages>0&&maxPages<=500&&Number.isSafeInteger(maxClusters)&&maxClusters>0&&maxClusters<=100000,'invalid','Invalid layout budget');
    const signature=JSON.stringify([provider.environmentId,options.paragraphShaper?.environmentId,maxPages,maxClusters,snapshot.document.documentId,model.defaultParagraphStyleId]);
    if(signature!==this.signature){this.entries.clear();this.signature=signature;}
    const next=new Map<string,Entry>(),pages:DeepReadonly<WriterPage>[]=[];const warnings=new Set<string>();
    let clusters=0,blocks=0;this.stats={sectionsComputed:0,sectionsReused:0,pagesReused:0};
    for(const section of model.sections){
      options.cancellation?.throwIfCancelled();
      const old=this.entries.get(section.id);
      let entry:Entry;
      if(old&&old.source===section&&old.defaults===model.defaultTextStyle&&old.styles===model.styles&&old.resources===model.resources){
        entry=old;this.stats.sectionsReused++;this.stats.pagesReused+=old.layout.pages.length;
      }else{
        const counts=cost(section);assert(clusters+counts.clusters<=maxClusters&&blocks+counts.blocks<=10000,'limit','Writer layout budget exceeded');
        assert(pages.length<maxPages,'limit','Writer page budget exceeded');
        const computed=layoutWriter({...snapshot,document:{...model,sections:[section]}},provider,{...options,maxPages:Math.max(1,maxPages-pages.length),maxClusters:Math.max(1,maxClusters-clusters)});
        const retained=computed.pages.map((p,i)=>{
          const prior=old?.layout.pages[i];
          // A page is reusable only when every geometry/semantic field agrees.
          if(prior&&JSON.stringify(prior)===JSON.stringify(p)){this.stats.pagesReused++;return prior;}return p;
        });
        entry={source:section,defaults:model.defaultTextStyle,styles:model.styles,resources:model.resources,...counts,layout:freezeData({...computed,pages:retained})};
        this.stats.sectionsComputed++;
      }
      clusters+=entry.clusters;blocks+=entry.blocks;
      assert(clusters<=maxClusters&&blocks<=10000&&pages.length+entry.layout.pages.length<=maxPages,'limit','Writer layout budget exceeded');
      pages.push(...entry.layout.pages);for(const w of entry.layout.warnings)warnings.add(w);next.set(section.id,entry);
    }
    options.cancellation?.throwIfCancelled();this.entries=next;
    return freezeData({revision:snapshot.revision,contentId:snapshot.contentId,environmentId:options.paragraphShaper?.environmentId??provider.environmentId,pages,warnings:[...warnings].sort()});
  }
}

