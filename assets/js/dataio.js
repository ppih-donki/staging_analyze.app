// assets/js/dataio.js
import { Papa } from '../vendor/globals.js';

export async function parseCSVFile(file, config={header:true, skipEmptyLines:true}){
  return new Promise((resolve, reject)=>{
    Papa.parse(file, { ...config, complete:res=>resolve(res), error:err=>reject(err) });
  });
}
