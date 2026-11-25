// SPDX-License-Identifier: MIT
/** One logical paragraph projection for serialization, editor and scene. */
export function splitSlideParagraphs<T extends {text:string}>(runs:readonly T[]):T[][]{
  const paragraphs:T[][]=[[]];
  for(const run of runs){
    const parts=run.text.split('\n');
    for(let i=0;i<parts.length;i++){
      if(i)paragraphs.push([]);
      paragraphs.at(-1)!.push({...run,text:parts[i]!});
    }
  }
  return paragraphs;
}
export function firstSlideParagraphRun<T extends {text:string}>(runs:readonly T[]):T|undefined{
  return runs.find(run=>run.text.length>0)??runs[0];
}
