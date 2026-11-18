// Bundle the results viewer into viewer.js and style.css, with icons and fonts
// embedded so the page works from file:// with no network access.
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import * as esbuild from 'esbuild';
const require=createRequire(import.meta.url);

const dataUris=(directory,type)=>Object.fromEntries(fs.readdirSync(directory).filter(file=>!file.startsWith('.'))
  .map(file=>[file,`data:${type(file)};base64,${fs.readFileSync(path.join(directory,file)).toString('base64')}`]));

export async function bundleInspector({sourceRoot,output}) {
  const icons=dataUris(path.join(sourceRoot,'viewer/icons'),file=>`image/${file.endsWith('.png')?'png':'webp'}`);
  const websiteAssets=path.join(sourceRoot,'viewer/website/assets');
  const websiteIcons=dataUris(path.join(websiteAssets,'icons'),()=>'image/svg+xml');
  let websiteCss='';
  const viewerFiles={name:'viewer-files',setup(build){
    // CSS modules: scoped class names for the components, styles collected into website.css.
    build.onLoad({filter:/\.module\.css$/},args=>{
      const {transform}=require('lightningcss');
      const result=transform({filename:args.path,code:fs.readFileSync(args.path),cssModules:{pattern:'abw_[local]'},targets:{chrome:120<<16},minify:true});
      websiteCss+=result.code.toString();
      return {contents:`export default ${JSON.stringify(Object.fromEntries(Object.entries(result.exports).map(([key,value])=>[key,value.name])))};`,loader:'js'};
    });
    build.onLoad({filter:/\.css$/},args=>({contents:fs.readFileSync(args.path,'utf8'),loader:'text'}));
    // Product icons are referenced by their website URL; embed them instead.
    build.onLoad({filter:/\.[cm]?jsx?$/},args=>{
      if(args.path.includes(`${path.sep}node_modules${path.sep}`))return undefined;
      let source=fs.readFileSync(args.path,'utf8');
      for(const [file,uri] of Object.entries(icons))source=source.replaceAll(`"/benchmark-icons/${file}"`,JSON.stringify(uri));
      for(const [file,uri] of Object.entries(websiteIcons))source=source.replaceAll(`"/icons/${file}"`,JSON.stringify(uri));
      return {contents:source,loader:'jsx'};
    });
  }};
  await esbuild.build({entryPoints:[path.join(sourceRoot,'viewer/Viewer.jsx')],bundle:true,
    outfile:path.join(output,'viewer.js'),format:'iife',target:['chrome120'],minify:true,jsx:'automatic',
    define:{'process.env.NODE_ENV':'"production"'},plugins:[viewerFiles],logLevel:'silent'});
  fs.writeFileSync(path.join(output,'website.css'),websiteCss);
  const postcss=require('postcss'),tailwind=require('@tailwindcss/postcss');
  const css=await postcss([tailwind({base:sourceRoot,optimize:true})]).process(`@import "tailwindcss" source(none);\n@source "${path.join(sourceRoot,'viewer/components').replaceAll('\\','/')}";\n@source "${path.join(sourceRoot,'viewer').replaceAll('\\','/')}";\n`,{from:path.join(sourceRoot,'viewer/build.css')});
  const fonts=[['Inter','Inter-VariableFont_opsz,wght.woff2'],['Lexend Deca','LexendDeca-VariableFont_wght.woff2']].map(([name,file])=>`@font-face{font-family:"${name}";font-style:normal;font-weight:100 900;font-display:swap;src:url(data:font/woff2;base64,${fs.readFileSync(path.join(websiteAssets,'fonts',file)).toString('base64')}) format("woff2");}`).join('\n');
  fs.writeFileSync(path.join(output,'style.css'),css.css+'\n'+fs.readFileSync(path.join(sourceRoot,'viewer/components/theme.css'),'utf8')+'\n'+fonts+'\n'+fs.readFileSync(path.join(sourceRoot,'viewer/website/tokens.css'),'utf8')+'\n'+websiteCss+'\n'+fs.readFileSync(path.join(sourceRoot,'viewer/style.css'),'utf8'));
  return {renderer:'website-charts-and-shared-inspector/v2',assetsEmbedded:true};
}
