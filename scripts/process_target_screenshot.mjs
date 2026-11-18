// Pixel processing only. No browser, screenshot source, or workbook access.
import fs from 'node:fs';
import {processTargetScreenshot} from './capture_targeted_sheets.mjs';
const [input,output,options]=process.argv.slice(2);
if(!input||!output||!options)throw Error('Expected input, output and measured crop options');
const {cropped,...result}=await processTargetScreenshot(fs.readFileSync(input),JSON.parse(options),{hostProcess:true});
fs.writeFileSync(output,cropped,{flag:'wx'});
process.stdout.write(JSON.stringify(result));
