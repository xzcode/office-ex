// SPDX-License-Identifier: MIT
/** Test helpers navigate real tabs and panels rather than forcing hidden DOM
 * controls visible. Missing selectors still fail in the calling assertion. */
export async function revealControl(page,selector){
 const target=page.locator(selector).first();
 const textMode=selector==='#draft-text'||selector.startsWith('.writer-edit-table')||selector.startsWith('.writer-image-card');
 if(textMode&&await target.count()===0)await page.waitForFunction(()=>document.querySelector('#draft-text,#writer-exit-page-edit'),null,{timeout:15000});
 if(await target.count()===0&&textMode&&await page.locator('#writer-exit-page-edit').count()){
   await page.locator('.office-ribbon-tab[data-ribbon-tab=view]').click();
   await page.locator('#writer-exit-page-edit').click();
   await page.locator('#draft-text').waitFor();
 }
 if(await target.count()===0)return page.locator(selector);
 const detailClasses=await target.evaluate(el=>{const out=[];for(let d=el.parentElement;d;d=d.parentElement)if(d instanceof HTMLDetailsElement&&!d.open){if(d.classList.contains('home-other-formats'))out.push('home-other-formats');else if(d.classList.contains('internal-drafts'))out.push('internal-drafts');}return out.reverse();});
 for(const cls of detailClasses){const details=page.locator('.'+cls);if(!await details.evaluate(el=>el.open))await details.locator(':scope > summary').click();}

 const panel=await target.evaluate(el=>el.closest('[data-ribbon-panel]')?.dataset.ribbonPanel);
 if(panel){const tab=page.locator('.office-ribbon-tab[data-ribbon-tab="'+panel+'"]');if(await tab.getAttribute('aria-selected')!=='true')await tab.click();}
 const property=await target.evaluate(el=>el.closest('fieldset[data-property-panel]')?.dataset.propertyPanel);
 if(property){const tab=page.locator('.office-properties-tabs [data-property-tab="'+property+'"]');if(await tab.getAttribute('aria-pressed')!=='true')await tab.click();}
 return page.locator(selector);
}
