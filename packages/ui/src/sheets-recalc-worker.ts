import {asFailure} from '@office-ex/contracts';
import {WorkbookCalculator,applyCalculationUpdate,type CalculationVersion,type CalculationUpdate,type CalculatedValue} from '@office-ex/sheets';
interface Job {id:number;revision:number;update:CalculationUpdate;cachedValues?:[string,CalculatedValue][];cachedSources?:[string,string][]}
const calculator=new WorkbookCalculator();
let previous:CalculationVersion|undefined,cachedValues=new Map<string,CalculatedValue>(),cachedSources=new Map<string,string>();
self.onmessage=(event:MessageEvent<Job>)=>{
 const job=event.data;
 try{
  const candidate=applyCalculationUpdate(previous,job.update);
  if(candidate.revision!==job.revision)throw new Error('Calculation revision mismatch');
  if(job.update.kind==='full'||job.cachedValues!==undefined)cachedValues=new Map(job.cachedValues??[]);
  if(job.update.kind==='full'||job.cachedSources!==undefined)cachedSources=new Map(job.cachedSources??[]);
  const result=calculator.calculate(candidate.document,{maxSteps:2_000_000,maxDepth:128,now:()=>performance.now(),deadline:performance.now()+10_000,cachedValues,cachedSources});
  previous=candidate;
  self.postMessage({id:job.id,revision:job.revision,ok:true,cells:[...result.cells],computedCells:result.computedCells,steps:result.steps});
 }catch(error){previous=undefined;calculator.clear();self.postMessage({id:job.id,revision:job.revision,ok:false,error:asFailure(error)});}
};
