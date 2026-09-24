// Генератор презентації (JavaScript, Node.js + pptxgenjs).
// Запуск: npm install pptxgenjs, потім node build_presentation.js
const pptxgen = require('pptxgenjs');
const p = new pptxgen(); p.layout = 'LAYOUT_16x9';
const D='1B2A41', M='324A5F', L='F4F6FA', A='E09F1F', W='FFFFFF', G='5B6B7C', H='Cambria', B='Calibri';
const sh = () => ({type:'outer', color:'000000', opacity:0.12, blur:6, offset:2, angle:90});
const title = (s, t, c) => s.addText(t, {x:0.5,y:0.35,w:9,h:0.7,fontFace:H,fontSize:28,bold:true,color:c,margin:0,valign:'middle',isTextBox:true});
const circle = (s, x, y, d, n, fs) => { s.addShape(p.ShapeType.ellipse,{x,y,w:d,h:d,fill:{color:A},line:{color:A}});
  s.addText(String(n),{x,y,w:d,h:d,align:'center',valign:'middle',fontFace:H,fontSize:fs,bold:true,color:D,margin:0,isTextBox:true}); };

// 1. Титул
let s = p.addSlide(); s.background = {color:D};
s.addText('ВАРІАНТ 1', {x:0.6,y:1.2,w:3,h:0.4,fontFace:B,fontSize:14,bold:true,color:A,charSpacing:4,margin:0,isTextBox:true});
s.addText('Класифікація документів', {x:0.6,y:1.7,w:6,h:1.5,fontFace:H,fontSize:40,bold:true,color:W,margin:0,valign:'top',isTextBox:true});
s.addText('Автобіографія', {x:0.6,y:3.4,w:6,h:0.6,fontFace:H,fontSize:26,color:A,margin:0,isTextBox:true});
s.addShape(p.ShapeType.ellipse,{x:7.0,y:1.5,w:2.4,h:2.4,fill:{color:M},line:{color:A,width:3}});
s.addText('1',{x:7.0,y:1.5,w:2.4,h:2.4,align:'center',valign:'middle',fontFace:H,fontSize:110,bold:true,color:A,margin:0,isTextBox:true});

// 2. Що таке автобіографія
s = p.addSlide(); s.background = {color:L}; title(s,'Що таке автобіографія',D);
s.addShape(p.ShapeType.roundRect,{x:0.5,y:1.3,w:4.4,h:3.6,rectRadius:0.1,fill:{color:W},line:{color:W},shadow:sh()});
s.addText([{text:'Автобіографія',options:{bold:true,color:M,breakLine:true}},
  {text:'документ особового характеру, у якому особа самостійно, від першої особи й у довільній формі викладає основні факти свого життя та діяльності.',options:{color:D}}],
  {x:0.8,y:1.5,w:3.8,h:3.2,fontFace:B,fontSize:18,valign:'middle',paraSpaceAfter:8,isTextBox:true});
[['Хто складає','Фізична особа власноруч'],['Кому подається','Роботодавцю, навчальному закладу, комісії'],['Коли потрібна','При прийомі на роботу, вступі, призначенні на посаду']]
 .forEach((r,i)=>{ const y=1.3+i*1.25; circle(s,5.3,y+0.2,0.6,i+1,20);
  s.addText(r[0],{x:6.1,y:y,w:3.4,h:0.4,fontFace:H,fontSize:17,bold:true,color:D,margin:0,valign:'bottom',isTextBox:true});
  s.addText(r[1],{x:6.1,y:y+0.42,w:3.4,h:0.7,fontFace:B,fontSize:14,color:G,margin:0,valign:'top',isTextBox:true}); });

// 3-4. Таблиці класифікації
const hdr = ['Ознака класифікації','Вид документа','Пояснення'].map(t=>({text:t,options:{bold:true,color:W,fill:{color:D},fontSize:13}}));
const tbl = (rows) => [hdr].concat(rows.map((r,i)=>[
  {text:r[0],options:{bold:true,color:M,fill:{color:i%2?'E9EDF3':W}}},
  {text:r[1],options:{bold:true,color:D,fill:{color:i%2?'E9EDF3':W}}},
  {text:r[2],options:{color:G,fill:{color:i%2?'E9EDF3':W}}}]));
const addTbl = (sl, rows) => sl.addTable(tbl(rows),{x:0.5,y:1.3,w:9,colW:[2.0,2.9,4.1],rowH:[0.4,0.66,0.66,0.66,0.66,0.66],fontFace:B,fontSize:12,valign:'middle',border:{type:'solid',pt:0.75,color:'D9DEE7'}});
s = p.addSlide(); s.background = {color:L}; title(s,'Класифікація автобіографії (1/2)',D);
addTbl(s,[
 ['За назвою','Автобіографія','Назва — обов’язковий реквізит, що вказує на вид документа'],
 ['За походженням','Особистий','Складає фізична особа від власного імені, а не установа'],
 ['За змістом','З особового складу (кадровий)','Входить до особової справи працівника'],
 ['За напрямом руху','Вхідний (для установи-одержувача)','Особа подає її при прийомі на роботу чи вступі'],
 ['За формою','Індивідуальний','Уніфікованого бланка немає, текст довільний, але зі сталою структурою']]);
s = p.addSlide(); s.background = {color:L}; title(s,'Класифікація автобіографії (2/2)',D);
addTbl(s,[
 ['За способом фіксації','Письмовий','Традиційно рукописна, допускається друкований варіант з підписом'],
 ['За стадією створення','Оригінал','Один примірник із власноручним підписом і датою'],
 ['За складністю','Простий','Розкриває одне питання — біографічні відомості про особу'],
 ['За ступенем доступу','З обмеженим доступом','Містить персональні дані, які захищені законом'],
 ['За терміном зберігання','Тривалого (75 років)','Зберігається у складі особової справи']]);

// 5. Структура
s = p.addSlide(); s.background = {color:L}; title(s,'Структура автобіографії',D);
const items = [['Назва','Слово «АВТОБІОГРАФІЯ» по центру'],['ПІБ','Повністю, від першої особи'],['Народження','Дата та місце народження'],['Освіта','Заклади, роки, спеціальність'],
 ['Трудова діяльність','Хронологічно, із зазначенням посад'],['Сім’я','Сімейний стан і склад сім’ї'],['Адреса','Місце проживання, контакти'],['Дата і підпис','Особистий підпис автора']];
items.forEach((it,i)=>{ const x=0.5+(i%4)*2.3125, y=1.3+Math.floor(i/4)*1.85;
  s.addShape(p.ShapeType.roundRect,{x,y,w:2.0625,h:1.6,rectRadius:0.1,fill:{color:W},line:{color:W},shadow:sh()});
  circle(s,x+0.15,y+0.15,0.42,i+1,14);
  s.addText(it[0],{x:x+0.15,y:y+0.65,w:1.8,h:0.3,fontFace:H,fontSize:13,bold:true,color:D,margin:0,isTextBox:true});
  s.addText(it[1],{x:x+0.15,y:y+0.97,w:1.8,h:0.55,fontFace:B,fontSize:11,color:G,margin:0,valign:'top',isTextBox:true}); });

// 6. Висновки
s = p.addSlide(); s.background = {color:D}; title(s,'Висновки',W);
s.addText('Автобіографія — особистий документ з особового складу, який має чітку структуру, підлягає обліку та зберігається в особовій справі.',
  {x:0.5,y:1.4,w:4.3,h:3.2,fontFace:B,fontSize:20,color:W,margin:0,valign:'top',isTextBox:true});
['Особистий','Вхідний','Письмовий','Індивідуальний','Простий','Оригінал','Обмеженого доступу','Тривалого зберігання'].forEach((t,i)=>{
  const x=5.1+(i%2)*2.35, y=1.4+Math.floor(i/2)*0.85;
  s.addShape(p.ShapeType.roundRect,{x,y,w:2.15,h:0.65,rectRadius:0.1,fill:{color:M},line:{color:A,width:1}});
  s.addText(t,{x,y,w:2.15,h:0.65,align:'center',valign:'middle',fontFace:B,fontSize:13,bold:true,color:W,margin:0,isTextBox:true}); });

p.writeFile({fileName:'Klasyfikatsiia_dokumentiv_variant1_avtobiohrafiia.pptx'}).then(()=>console.log('ok'));