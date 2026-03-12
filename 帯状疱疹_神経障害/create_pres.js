const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Medical Education';
const NAVY='1A3C5E', MID_BLUE='2980B9', LIGHT_BLUE='EBF5FB', ICE_BLUE='D6EAF8';
const WHITE='FFFFFF', OFF_WHITE='F7F9FC', DARK='1E293B', GRAY='94A3B8';
const RED='DC2626', GREEN='059669', YELLOW='D97706', TEAL='0D9488', FONT='Meiryo';
const mkS = () => ({ type:'outer', blur:8, offset:3, angle:135, color:'000000', opacity:0.08 });
const NL = "\n";
function slideFooter(s, num, total) {
  s.addShape(pres.shapes.RECTANGLE, {x:0, y:5.35, w:10, h:0.275, fill:{color:NAVY}});
  s.addText(num+' / '+total, {x:8.5, y:5.37, w:1.2, h:0.23, fontSize:10, fontFace:FONT, color:ICE_BLUE, align:'right', valign:'middle', margin:0});
  s.addText("医知創造ラボ", {x:0.3, y:5.37, w:2, h:0.23, fontSize:10, fontFace:FONT, color:ICE_BLUE, align:'left', valign:'middle', margin:0});
}
// S1 Title
let s1=pres.addSlide(); s1.background={color:NAVY};
s1.addText([{text:"\u5e2f\u72b6\u75b1\u75b9\u306f",options:{breakLine:true}},{text:"\u300c\u76ae\u819a\u306e\u75c5\u6c17\u300d\u3067\u306f\u306a\u3044",options:{}}],{x:0.8,y:1.0,w:8.4,h:2.2,fontSize:42,fontFace:FONT,color:WHITE,bold:true,align:'center',valign:'middle',lineSpacingMultiple:1.3});
s1.addShape(pres.shapes.RECTANGLE,{x:3.5,y:3.3,w:3,h:0.04,fill:{color:MID_BLUE}});
s1.addText("\u77e5\u3063\u3066\u304a\u304f\u3079\u304d\u795e\u7d4c\u969c\u5bb3\u306e\u5168\u4f53\u50cf",{x:1.5,y:3.6,w:7,h:0.8,fontSize:26,fontFace:FONT,color:ICE_BLUE,align:'center',valign:'middle'});
s1.addText("\u533b\u77e5\u5275\u9020\u30e9\u30dc",{x:0.5,y:4.8,w:9,h:0.5,fontSize:20,fontFace:FONT,color:GRAY,align:'center',valign:'middle'});
// S2 Hook
let s2=pres.addSlide(); s2.background={color:OFF_WHITE}; slideFooter(s2,2,19);
s2.addText("\u76ae\u75b9\u3068\u75db\u307f\u2026\u3067\u7d42\u308f\u3063\u3066\u3044\u307e\u305b\u3093\u304b\uff1f",{x:0.8,y:0.5,w:8.4,h:1.0,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center',valign:'middle'});
s2.addShape(pres.shapes.RECTANGLE,{x:1.2,y:1.8,w:7.6,h:2.8,fill:{color:WHITE},shadow:mkS()});
s2.addShape(pres.shapes.RECTANGLE,{x:1.2,y:1.8,w:0.08,h:2.8,fill:{color:RED}});
s2.addText([
  {text:"\u5e2f\u72b6\u75b1\u75b9\u306f\u5358\u306a\u308b\u76ae\u819a\u75be\u60a3\u3067\u306f\u306a\u304f",options:{breakLine:true,fontSize:24,fontFace:FONT,color:DARK}},
  {text:"\u300c\u795e\u7d4c\u306e\u75c5\u6c17\u300d\u3067\u3059",options:{breakLine:true,fontSize:28,fontFace:FONT,color:RED,bold:true}},
  {text:"",options:{breakLine:true,fontSize:20}},
  {text:"\u904b\u52d5\u9ebb\u75fa\u30fb\u6392\u5c3f\u969c\u5bb3\u30fb\u9854\u9762\u795e\u7d4c\u9ebb\u75fa\u2026",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:1.6,y:2.0,w:6.8,h:2.4,valign:'middle'});
// S3 TOC
let s3=pres.addSlide(); s3.background={color:OFF_WHITE}; slideFooter(s3,3,19);
s3.addText("\u672c\u65e5\u306e\u5185\u5bb9",{x:0.8,y:0.4,w:8.4,h:0.8,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
["\u5e2f\u72b6\u75b1\u75b9\u3068\u795e\u7d4c\u969c\u5bb3\u306e\u5168\u4f53\u50cf","\u611f\u899a\u795e\u7d4c\u969c\u5bb3\u3068PHN","\u904b\u52d5\u795e\u7d4c\u969c\u5bb3\uff08Zoster paresis\uff09","\u691c\u67fb\u3068\u6cbb\u7642","\u307e\u3068\u3081 \u2015 5\u3064\u306e\u30dd\u30a4\u30f3\u30c8"].forEach((item,i)=>{
  const yy=1.5+i*0.75;
  s3.addShape(pres.shapes.OVAL,{x:1.5,y:yy+0.05,w:0.5,h:0.5,fill:{color:MID_BLUE}});
  s3.addText(String(i+1),{x:1.5,y:yy+0.05,w:0.5,h:0.5,fontSize:22,fontFace:FONT,color:WHITE,align:'center',valign:'middle',margin:0});
  s3.addText(item,{x:2.2,y:yy,w:6,h:0.6,fontSize:24,fontFace:FONT,color:DARK,valign:'middle',margin:0});
});
// S4
let s4=pres.addSlide(); s4.background={color:OFF_WHITE}; slideFooter(s4,4,19);
s4.addText("\u5e2f\u72b6\u75b1\u75b9\u306e\u672c\u614b\u306f\u300c\u795e\u7d4c\u7bc0\u708e\u300d",{x:0.8,y:0.4,w:8.4,h:0.8,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
s4.addShape(pres.shapes.RECTANGLE,{x:0.8,y:1.5,w:8.4,h:3.2,fill:{color:WHITE},shadow:mkS()});
s4.addText([
  {text:"\u6c34\u75d8\u30fb\u5e2f\u72b6\u75b1\u75b9\u30a6\u30a4\u30eb\u30b9\uff08VZV\uff09\u306f",options:{breakLine:true,fontSize:24,fontFace:FONT,color:DARK}},
  {text:"\u5f8c\u6839\u795e\u7d4c\u7bc0\u306b\u6f5c\u4f0f\u611f\u67d3",options:{breakLine:true,fontSize:26,fontFace:FONT,color:MID_BLUE,bold:true}},
  {text:"",options:{breakLine:true,fontSize:20}},
  {text:"\u514d\u75ab\u4f4e\u4e0b\u6642\u306b\u30a6\u30a4\u30eb\u30b9\u304c\u518d\u6d3b\u6027\u5316\u3057",options:{breakLine:true,fontSize:24,fontFace:FONT,color:DARK}},
  {text:"\u795e\u7d4c\u306b\u6cbf\u3063\u3066\u708e\u75c7\u304c\u6ce2\u53ca\u3059\u308b",options:{breakLine:true,fontSize:24,fontFace:FONT,color:DARK}},
  {text:"",options:{breakLine:true,fontSize:20}},
  {text:"\u76ae\u75b9\u306f\u300c\u7d50\u679c\u300d\u3067\u3042\u308a\u3001\u672c\u614b\u306f\u795e\u7d4c\u708e",options:{fontSize:24,fontFace:FONT,color:RED,bold:true}}
],{x:1.2,y:1.7,w:7.6,h:2.8,valign:'middle',align:'center'});
// S5 Table
let s5=pres.addSlide(); s5.background={color:OFF_WHITE}; slideFooter(s5,5,19);
s5.addText("\u795e\u7d4c\u969c\u5bb3\u306e\u5168\u4f53\u50cf",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const H=(t)=>({text:t,options:{bold:true,color:WHITE,fill:{color:NAVY},fontSize:20,fontFace:FONT,align:'center',valign:'middle'}});
const CW=(t)=>({text:t,options:{fontSize:20,fontFace:FONT,fill:{color:WHITE},align:'center',valign:'middle'}});
const CB=(t)=>({text:t,options:{fontSize:20,fontFace:FONT,fill:{color:LIGHT_BLUE},align:'center',valign:'middle'}});
s5.addTable([
  [H("\u969c\u5bb3\u90e8\u4f4d"),H("\u4e3b\u306a\u75c7\u72b6")],
  [CW("\u5f8c\u6839\u795e\u7d4c\u7bc0"),CW("\u76ae\u75b9\u30fb\u795e\u7d4c\u75db")],
  [CB("\u611f\u899a\u795e\u7d4c"),CB("\u611f\u899a\u969c\u5bb3")],
  [CW("\u524d\u89d2\u30fb\u904b\u52d5\u795e\u7d4c"),CW("\u904b\u52d5\u9ebb\u75fa")],
  [CB("\u8133\u795e\u7d4c"),CB("\u9854\u9762\u795e\u7d4c\u9ebb\u75fa \u7b49")],
  [CW("\u81ea\u5f8b\u795e\u7d4c"),CW("\u6392\u5c3f\u969c\u5bb3")]
],{x:1.0,y:1.2,w:8.0,colW:[3.5,4.5],border:{pt:1,color:'D5D8DC'},rowH:[0.55,0.55,0.55,0.55,0.55,0.55]});
s5.addText("\u5e2f\u72b6\u75b1\u75b9\u306f\u591a\u5f69\u306a\u795e\u7d4c\u969c\u5bb3\u3092\u5f15\u304d\u8d77\u3053\u3057\u3046\u308b",{x:1.0,y:4.7,w:8.0,h:0.5,fontSize:20,fontFace:FONT,color:MID_BLUE,italic:true,align:'center'});
// S6 Three important disorders
let s6=pres.addSlide(); s6.background={color:OFF_WHITE}; slideFooter(s6,6,19);
s6.addText("\u81e8\u5e8a\u3067\u7279\u306b\u91cd\u8981\u306a3\u3064\u306e\u969c\u5bb3",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const cards6=[
  {t:"\u5e2f\u72b6\u75b1\u75b9\u5f8c\u795e\u7d4c\u75db",sub:"PHN",c:RED},
  {t:"\u904b\u52d5\u9ebb\u75fa",sub:"Zoster paresis",c:MID_BLUE},
  {t:"\u8133\u795e\u7d4c\u969c\u5bb3",sub:"\u9854\u9762\u795e\u7d4c\u9ebb\u75fa\u7b49",c:GREEN}
];
cards6.forEach((cd,i)=>{
  const cx=0.6+i*3.1;
  s6.addShape(pres.shapes.RECTANGLE,{x:cx,y:1.3,w:2.9,h:3.0,fill:{color:WHITE},shadow:mkS()});
  s6.addShape(pres.shapes.RECTANGLE,{x:cx,y:1.3,w:2.9,h:0.08,fill:{color:cd.c}});
  s6.addText(cd.t,{x:cx+0.2,y:1.8,w:2.5,h:1.0,fontSize:24,fontFace:FONT,color:DARK,bold:true,align:'center',valign:'middle'});
  s6.addText(cd.sub,{x:cx+0.2,y:2.9,w:2.5,h:0.8,fontSize:20,fontFace:FONT,color:GRAY,align:'center',valign:'middle'});
});
// S7 Acute pain
let s7=pres.addSlide(); s7.background={color:OFF_WHITE}; slideFooter(s7,7,19);
s7.addText("\u6025\u6027\u671f\u306e\u795e\u7d4c\u75db",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const pains=[
  {t:"\u96fb\u6483\u75db",d:"\u7a81\u7136\u306e\u92ed\u3044\u75db\u307f\u304c\u8d70\u308b",c:RED},
  {t:"Burning pain",d:"\u7126\u3051\u308b\u3088\u3046\u306a\u6301\u7d9a\u7684\u306a\u75db\u307f",c:YELLOW},
  {t:"\u30a2\u30ed\u30c7\u30a3\u30cb\u30a2",d:"\u8efd\u3044\u63a5\u89e6\u3067\u8a98\u767a\u3055\u308c\u308b\u75db\u307f",c:MID_BLUE}
];
pains.forEach((p,i)=>{
  const py=1.3+i*1.3;
  s7.addShape(pres.shapes.RECTANGLE,{x:1.2,y:py,w:7.6,h:1.1,fill:{color:WHITE},shadow:mkS()});
  s7.addShape(pres.shapes.RECTANGLE,{x:1.2,y:py,w:0.08,h:1.1,fill:{color:p.c}});
  s7.addText(p.t,{x:1.6,y:py,w:2.5,h:1.1,fontSize:24,fontFace:FONT,color:DARK,bold:true,valign:'middle',margin:0});
  s7.addText(p.d,{x:4.2,y:py,w:4.2,h:1.1,fontSize:22,fontFace:FONT,color:DARK,valign:'middle',margin:0});
});
// S8 PHN definition
let s8=pres.addSlide(); s8.background={color:OFF_WHITE}; slideFooter(s8,8,19);
s8.addText("\u5e2f\u72b6\u75b1\u75b9\u5f8c\u795e\u7d4c\u75db\uff08PHN\uff09",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
s8.addShape(pres.shapes.RECTANGLE,{x:1.0,y:1.3,w:8.0,h:1.8,fill:{color:NAVY},shadow:mkS()});
s8.addText([
  {text:"\u76ae\u75b9\u6cbb\u7652\u5f8c 3\u304b\u6708\u4ee5\u4e0a",options:{breakLine:true,fontSize:28,fontFace:FONT,color:WHITE,bold:true}},
  {text:"\u7d9a\u304f\u75db\u307f",options:{fontSize:28,fontFace:FONT,color:ICE_BLUE,bold:true}}
],{x:1.4,y:1.5,w:7.2,h:1.4,align:'center',valign:'middle'});
s8.addShape(pres.shapes.RECTANGLE,{x:1.0,y:3.5,w:8.0,h:1.6,fill:{color:WHITE},shadow:mkS()});
s8.addText([
  {text:"\u6700\u3082\u983b\u5ea6\u304c\u9ad8\u304f\u3001QOL\u3092\u8457\u3057\u304f\u4f4e\u4e0b\u3055\u305b\u308b",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u5e2f\u72b6\u75b1\u75b9\u306e\u795e\u7d4c\u5408\u4f75\u75c7\u306e\u4e2d\u3067\u6700\u91cd\u8981",options:{fontSize:22,fontFace:FONT,color:RED,bold:true}}
],{x:1.4,y:3.7,w:7.2,h:1.2,align:'center',valign:'middle'});
// S9 PHN rates bar chart
let s9=pres.addSlide(); s9.background={color:OFF_WHITE}; slideFooter(s9,9,19);
s9.addText("PHN\u767a\u75c7\u7387\uff08\u5e74\u9f62\u5225\uff09",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
s9.addChart(pres.charts.BAR,[{
  name:"PHN\u767a\u75c7\u7387",
  labels:["50\u6b73\u672a\u6e80","60\u6b73\u4ee5\u4e0a","80\u6b73\u4ee5\u4e0a"],
  values:[5,15,30]
}],{x:1.0,y:1.2,w:8.0,h:3.5,barDir:'bar',
  chartColors:[MID_BLUE],
  showValue:true,dataLabelPosition:'outEnd',dataLabelColor:DARK,
  catAxisLabelColor:DARK,catAxisLabelFontSize:20,
  valAxisLabelColor:GRAY,valAxisLabelFontSize:20,
  valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},
  showLegend:false,
  chartArea:{fill:{color:WHITE},roundedCorners:true}
});
s9.addText("\u9ad8\u9f62\u8005\u307b\u3069PHN\u30ea\u30b9\u30af\u304c\u9ad8\u3044",{x:1.0,y:4.9,w:8.0,h:0.4,fontSize:20,fontFace:FONT,color:MID_BLUE,italic:true,align:'center'});
// S10 Risk factors and prognosis
let s10=pres.addSlide(); s10.background={color:OFF_WHITE}; slideFooter(s10,10,19);
s10.addText("PHN\u306e\u5371\u967a\u56e0\u5b50\u3068\u4e88\u5f8c",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
s10.addShape(pres.shapes.RECTANGLE,{x:0.6,y:1.2,w:4.2,h:3.6,fill:{color:WHITE},shadow:mkS()});
s10.addShape(pres.shapes.RECTANGLE,{x:0.6,y:1.2,w:4.2,h:0.08,fill:{color:RED}});
s10.addText("\u5371\u967a\u56e0\u5b50",{x:0.8,y:1.4,w:3.8,h:0.6,fontSize:24,fontFace:FONT,color:RED,bold:true,align:'center'});
s10.addText([
  {text:"\u2022 \u9ad8\u9f62",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u2022 \u6fc0\u3057\u3044\u6025\u6027\u671f\u75db",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u2022 \u76ae\u75b9\u7bc4\u56f2\u304c\u5e83\u3044",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u2022 \u514d\u75ab\u6291\u5236",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:0.8,y:2.1,w:3.8,h:2.5,valign:'middle'});
s10.addShape(pres.shapes.RECTANGLE,{x:5.2,y:1.2,w:4.2,h:3.6,fill:{color:WHITE},shadow:mkS()});
s10.addShape(pres.shapes.RECTANGLE,{x:5.2,y:1.2,w:4.2,h:0.08,fill:{color:GREEN}});
s10.addText("\u4e88\u5f8c",{x:5.4,y:1.4,w:3.8,h:0.6,fontSize:24,fontFace:FONT,color:GREEN,bold:true,align:'center'});
s10.addText([
  {text:"1\u5e74\u4ee5\u5185\u6539\u5584 \u2026 \u7d0450%",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u6570\u5e74\u6301\u7d9a \u2026 \u7d0420%",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u6c38\u7d9a \u2026 \u5c11\u6570",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:5.4,y:2.1,w:3.8,h:2.5,valign:'middle'});
// S11 Motor nerve section divider
let s11=pres.addSlide(); s11.background={color:NAVY};
s11.addText([
  {text:"\u904b\u52d5\u795e\u7d4c\u969c\u5bb3",options:{breakLine:true,fontSize:38,fontFace:FONT,color:WHITE,bold:true}},
  {text:"Zoster paresis",options:{fontSize:24,fontFace:FONT,color:ICE_BLUE}}
],{x:0.8,y:0.8,w:8.4,h:1.5,align:'center',valign:'middle'});
s11.addText("3\uff5e5%",{x:2.5,y:2.5,w:5,h:1.5,fontSize:72,fontFace:FONT,color:MID_BLUE,bold:true,align:'center',valign:'middle'});
s11.addText("\u306e\u5e2f\u72b6\u75b1\u75b9\u60a3\u8005\u306b\u904b\u52d5\u9ebb\u75fa\u304c\u51fa\u73fe",{x:1.0,y:4.0,w:8.0,h:0.8,fontSize:24,fontFace:FONT,color:ICE_BLUE,align:'center',valign:'middle'});
s11.addText("\u76ae\u75b9\u51fa\u73fe\u304b\u30892\uff5e3\u9031\u9593\u5f8c\u306b\u767a\u75c7",{x:1.0,y:4.7,w:8.0,h:0.5,fontSize:20,fontFace:FONT,color:GRAY,align:'center',valign:'middle'});
// S12 Segmental paralysis
let s12=pres.addSlide(); s12.background={color:OFF_WHITE}; slideFooter(s12,12,19);
s12.addText("\u56db\u80a2\u306e\u5206\u7bc0\u6027\u9ebb\u75fa",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const flow12=[
  {t:"\u7f6a\u60a3\u795e\u7d4c\u9ad8\u4f4d",c:NAVY},
  {t:"\u524d\u89d2\u7d30\u80de\u969c\u5bb3",c:MID_BLUE},
  {t:"\u5206\u7bc0\u6027\u904b\u52d5\u9ebb\u75fa",c:RED}
];
flow12.forEach((f,i)=>{
  const fx=0.8+i*3.2;
  s12.addShape(pres.shapes.RECTANGLE,{x:fx,y:1.3,w:2.8,h:1.2,fill:{color:f.c},shadow:mkS()});
  s12.addText(f.t,{x:fx,y:1.3,w:2.8,h:1.2,fontSize:22,fontFace:FONT,color:WHITE,bold:true,align:'center',valign:'middle'});
  if(i<2){
    s12.addShape(pres.shapes.LINE,{x:fx+2.8,y:1.9,w:0.4,h:0,line:{color:GRAY,width:2}});
  }
});
s12.addShape(pres.shapes.RECTANGLE,{x:0.8,y:3.0,w:8.4,h:2.0,fill:{color:WHITE},shadow:mkS()});
s12.addText([
  {text:"上肢：C5\uff5eT1領域 → 上肢の挙上が困難",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u4e0b\u80a2\uff1aL2\uff5eS1\u9818\u57df \u2192 \u4e0b\u80a2\u306e\u8131\u529b",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"\u4f53\u5e79\uff1a\u8179\u7b4b\u9ebb\u75fa \u2192 \u8179\u58c1\u81a8\u9686",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:1.2,y:3.2,w:7.6,h:1.6,valign:'middle'});
// S13 Ramsay Hunt
let s13=pres.addSlide(); s13.background={color:OFF_WHITE}; slideFooter(s13,13,19);
s13.addText("\u30e9\u30e0\u30bc\u30a4\u30fb\u30cf\u30f3\u30c8\u75c7\u5019\u7fa4",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const hunt=[
  {t:"\u8033\u4ecb\u30fb\u5916\u8033\u9053\u306e\u75b1\u75b9",c:RED},
  {t:"\u9854\u9762\u795e\u7d4c\u9ebb\u75fa",c:MID_BLUE},
  {t:"\u8074\u899a\u969c\u5bb3",c:YELLOW},
  {t:"\u3081\u307e\u3044\u30fb\u5e73\u8861\u969c\u5bb3",c:GREEN}
];
hunt.forEach((h,i)=>{
  const row=Math.floor(i/2), col=i%2;
  const hx=0.8+col*4.4, hy=1.3+row*1.8;
  s13.addShape(pres.shapes.RECTANGLE,{x:hx,y:hy,w:4.0,h:1.4,fill:{color:WHITE},shadow:mkS()});
  s13.addShape(pres.shapes.RECTANGLE,{x:hx,y:hy,w:0.08,h:1.4,fill:{color:h.c}});
  s13.addText(h.t,{x:hx+0.3,y:hy,w:3.5,h:1.4,fontSize:24,fontFace:FONT,color:DARK,bold:true,valign:'middle',margin:0});
});
s13.addText("\u7b2c7\u8133\u795e\u7d4c\u306e\u819d\u795e\u7d4c\u7bc0\u306bVZV\u304c\u6f5c\u4f0f",{x:1.0,y:4.8,w:8.0,h:0.5,fontSize:20,fontFace:FONT,color:MID_BLUE,italic:true,align:'center'});
// S14 Other motor disorders
let s14=pres.addSlide(); s14.background={color:OFF_WHITE}; slideFooter(s14,14,19);
s14.addText("\u305d\u306e\u4ed6\u306e\u904b\u52d5\u969c\u5bb3",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
s14.addShape(pres.shapes.RECTANGLE,{x:0.6,y:1.2,w:4.2,h:3.6,fill:{color:WHITE},shadow:mkS()});
s14.addShape(pres.shapes.RECTANGLE,{x:0.6,y:1.2,w:4.2,h:0.08,fill:{color:RED}});
s14.addText("横隔神経麻痺",{x:0.8,y:1.5,w:3.8,h:0.6,fontSize:24,fontFace:FONT,color:RED,bold:true,align:'center'});
s14.addText([
  {text:"C3〜C5領域の帯状疱疹",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"横隔神経麻痺 → 呼吸困難",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"胸部X線で横隔膜挙上を確認",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:0.8,y:2.3,w:3.8,h:2.2,valign:'middle'});
s14.addShape(pres.shapes.RECTANGLE,{x:5.2,y:1.2,w:4.2,h:3.6,fill:{color:WHITE},shadow:mkS()});
s14.addShape(pres.shapes.RECTANGLE,{x:5.2,y:1.2,w:4.2,h:0.08,fill:{color:MID_BLUE}});
s14.addText("膀胱直腸障害",{x:5.4,y:1.5,w:3.8,h:0.6,fontSize:24,fontFace:FONT,color:MID_BLUE,bold:true,align:'center'});
s14.addText([
  {text:"S2〜S4領域の帯状疱疹",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"排尿障害・尿閉",options:{breakLine:true,fontSize:22,fontFace:FONT,color:DARK}},
  {text:"多くは一過性だが注意必要",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:5.4,y:2.3,w:3.8,h:2.2,valign:'middle'});
// S15 Motor paralysis prognosis
let s15=pres.addSlide(); s15.background={color:OFF_WHITE}; slideFooter(s15,15,19);
s15.addText("\u904b\u52d5\u9ebb\u75fa\u306e\u4e88\u5f8c",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const progs=[
  {t:"\u5b8c\u5168\u56de\u5fa9",v:"50\uff5e70%",c:GREEN},
  {t:"\u90e8\u5206\u56de\u5fa9",v:"20\uff5e30%",c:YELLOW},
  {t:"\u5f8c\u907a\u75c7",v:"\u7d0410%",c:RED}
];
progs.forEach((p,i)=>{
  const px=0.6+i*3.2;
  s15.addShape(pres.shapes.RECTANGLE,{x:px,y:1.2,w:2.9,h:2.0,fill:{color:WHITE},shadow:mkS()});
  s15.addText(p.v,{x:px,y:1.3,w:2.9,h:1.0,fontSize:36,fontFace:FONT,color:p.c,bold:true,align:'center',valign:'middle'});
  s15.addText(p.t,{x:px,y:2.2,w:2.9,h:0.8,fontSize:22,fontFace:FONT,color:DARK,align:'center',valign:'middle'});
});
s15.addShape(pres.shapes.RECTANGLE,{x:0.6,y:3.6,w:8.8,h:1.6,fill:{color:WHITE},shadow:mkS()});
s15.addShape(pres.shapes.RECTANGLE,{x:0.6,y:3.6,w:0.08,h:1.6,fill:{color:MID_BLUE}});
s15.addText([
  {text:"\u56de\u5fa9\u671f\u9593\u306e\u76ee\u5b89",options:{breakLine:true,fontSize:22,fontFace:FONT,color:MID_BLUE,bold:true}},
  {text:"\u8efd\u75c7\uff1a\u7d043\u304b\u6708\u3000\u4e00\u822c\u7684\uff1a6\uff5e12\u304b\u6708\u3000\u91cd\u75c7\uff1a2\u5e74\u4ee5\u4e0a",options:{fontSize:22,fontFace:FONT,color:DARK}}
],{x:1.0,y:3.8,w:8.0,h:1.2,valign:'middle'});
// S16 Tests
let s16=pres.addSlide(); s16.background={color:OFF_WHITE}; slideFooter(s16,16,19);
s16.addText("\u691c\u67fb",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const TH=(t)=>({text:t,options:{bold:true,color:WHITE,fill:{color:NAVY},fontSize:20,fontFace:FONT,align:'center',valign:'middle'}});
const TCW=(t)=>({text:t,options:{fontSize:20,fontFace:FONT,fill:{color:WHITE},valign:'middle'}});
const TCB=(t)=>({text:t,options:{fontSize:20,fontFace:FONT,fill:{color:LIGHT_BLUE},valign:'middle'}});
s16.addTable([
  [TH("\u691c\u67fb"),TH("\u76ee\u7684"),TH("\u6240\u898b")],
  [TCW("MRI"),TCW("神経炎症の確認"),TCW("神経の腫大・造影効果")],
  [TCB("神経伝導検査"),TCB("運動障害の評価"),TCB("脱神経所見")],
  [TCW("髄液検査"),TCW("中枢神経障害"),TCW("細胞数増加・VZV-DNA")]
],{x:0.6,y:1.2,w:8.8,colW:[2.2,3.3,3.3],border:{pt:1,color:'D5D8DC'},rowH:[0.6,0.7,0.7,0.7]});
// S17 Treatment
let s17=pres.addSlide(); s17.background={color:OFF_WHITE}; slideFooter(s17,17,19);
s17.addText("\u6cbb\u7642",{x:0.8,y:0.3,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const treats=[
  {t:"\u6297\u30a6\u30a4\u30eb\u30b9\u85ac",d:"72\u6642\u9593\u4ee5\u5185\u306b\u958b\u59cb",sub:"\u30a2\u30b7\u30af\u30ed\u30d3\u30eb / \u30d0\u30e9\u30b7\u30af\u30ed\u30d3\u30eb",c:RED},
  {t:"神経痛治療",d:"急性期から積極的に",sub:"プレガバリン / ガバペンチン\n三環系抗うつ薬 / 神経ブロック",c:YELLOW},
  {t:"運動麻痺",d:"保存的治療＋リハビリ",sub:"顔面神経麻痺ではステロイド併用",c:GREEN}
];
treats.forEach((tr,i)=>{
  const tx=0.5+i*3.2;
  s17.addShape(pres.shapes.RECTANGLE,{x:tx,y:1.2,w:2.9,h:3.8,fill:{color:WHITE},shadow:mkS()});
  s17.addShape(pres.shapes.RECTANGLE,{x:tx,y:1.2,w:2.9,h:0.08,fill:{color:tr.c}});
  s17.addText(tr.t,{x:tx+0.2,y:1.5,w:2.5,h:0.8,fontSize:24,fontFace:FONT,color:DARK,bold:true,align:'center',valign:'middle'});
  s17.addText(tr.d,{x:tx+0.2,y:2.4,w:2.5,h:0.8,fontSize:20,fontFace:FONT,color:MID_BLUE,align:'center',valign:'middle'});
  s17.addText(tr.sub,{x:tx+0.2,y:3.4,w:2.5,h:1.2,fontSize:20,fontFace:FONT,color:GRAY,align:'center',valign:'middle'});
});
// S18 Summary 5 points
let s18=pres.addSlide(); s18.background={color:OFF_WHITE}; slideFooter(s18,18,19);
s18.addText("\u307e\u3068\u3081 \u2015 5\u3064\u306e\u30dd\u30a4\u30f3\u30c8",{x:0.8,y:0.2,w:8.4,h:0.7,fontSize:34,fontFace:FONT,color:NAVY,bold:true,align:'center'});
const sums=[
  "\u5e2f\u72b6\u75b1\u75b9\u306f\u300c\u795e\u7d4c\u306e\u75c5\u6c17\u300d\u3067\u3042\u308b",
  "PHN\u306f\u6700\u3082\u591a\u3044\u5408\u4f75\u75c7\uff08\u9ad8\u9f62\u8005\u306f30%\uff09",
  "3\uff5e5%\u306b\u904b\u52d5\u9ebb\u75fa\u304c\u8d77\u3053\u308a\u3046\u308b",
  "\u6297\u30a6\u30a4\u30eb\u30b9\u85ac\u306f72\u6642\u9593\u4ee5\u5185\u304c\u9375",
  "\u795e\u7d4c\u969c\u5bb3\u3092\u77e5\u308c\u3070\u65e9\u671f\u767a\u898b\u3067\u304d\u308b"
];
sums.forEach((s,i)=>{
  const sy=1.1+i*0.85;
  s18.addShape(pres.shapes.OVAL,{x:0.8,y:sy+0.05,w:0.5,h:0.5,fill:{color:MID_BLUE}});
  s18.addText(String(i+1),{x:0.8,y:sy+0.05,w:0.5,h:0.5,fontSize:22,fontFace:FONT,color:WHITE,align:'center',valign:'middle',margin:0});
  s18.addText(s,{x:1.5,y:sy,w:7.5,h:0.6,fontSize:22,fontFace:FONT,color:DARK,valign:'middle',margin:0});
});
// S19 End slide
let s19=pres.addSlide(); s19.background={color:NAVY};
s19.addText([
  {text:"\u3054\u8996\u8074\u3042\u308a\u304c\u3068\u3046\u3054\u3056\u3044\u307e\u3057\u305f",options:{breakLine:true,fontSize:36,fontFace:FONT,color:WHITE,bold:true}},
  {text:"",options:{breakLine:true,fontSize:20}},
  {text:"\u30c1\u30e3\u30f3\u30cd\u30eb\u767b\u9332\u30fb\u9ad8\u8a55\u4fa1\u304a\u9858\u3044\u3057\u307e\u3059",options:{fontSize:24,fontFace:FONT,color:ICE_BLUE}}
],{x:0.8,y:1.5,w:8.4,h:2.5,align:'center',valign:'middle'});
s19.addText("\u533b\u77e5\u5275\u9020\u30e9\u30dc",{x:0.5,y:4.5,w:9,h:0.5,fontSize:22,fontFace:FONT,color:GRAY,align:'center',valign:'middle'});
// Write file
pres.writeFile({fileName:"C:/Users/jsber/OneDrive/Documents/Claude_task/帯状疱疹_神経障害_プレゼン_v2.pptx"}).then(()=>console.log("PPTX created successfully!")).catch(e=>console.error(e));
