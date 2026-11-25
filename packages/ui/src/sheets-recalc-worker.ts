import { asFailure } from '@office-ex/contracts';
import { WorkbookCalculator, type WorkbookSnapshot, type CalculatedValue } from '@office-ex/sheets';
interface Job {id:number;revision:number;document:WorkbookSnapshot;cachedValues?:[string,CalculatedValue][];cachedSources?:[string,string][]}
const calculator=new WorkbookCalculator();
self.onmessage=(event:MessageEvent<Job>)=>{
  const job=event.data;
  try{
    const result=calculator.calculate(job.document,{maxSteps:2_000_000,maxDepth:128,now:()=>performance.now(),deadline:performance.now()+10_000,cachedValues:new Map(job.cachedValues??[]),cachedSources:new Map(job.cachedSources??[])});
    self.postMessage({id:job.id,revision:job.revision,ok:true,cells:[...result.cells],computedCells:result.computedCells,steps:result.steps});
  }catch(error){self.postMessage({id:job.id,revision:job.revision,ok:false,error:asFailure(error)});}
};
