// SPDX-License-Identifier: MIT
/** Older multi-tab regression scripts explicitly exercise the retained
 * optional tabbed mode. Default-window tests use an unseeded profile. */
import {mkdir,writeFile} from 'node:fs/promises';
import {join} from 'node:path';
export async function prepareTabbedProfile(path){
 await mkdir(path,{recursive:true});
 await writeFile(join(path,'settings.json'),JSON.stringify({version:2,theme:'system',language:'system',region:'system',documentLanguage:'system',openInTabs:true,recent:[]}),{flag:'wx'});
}
