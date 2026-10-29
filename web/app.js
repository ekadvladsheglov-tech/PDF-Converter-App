const S=[
 {id:'p2w',g:'Конвертация',ic:'🔄',t:'PDF → Word',sub:'Конвертация документов',mode:1,ftype:'pdf',prog:1,act:'run_p2w',actLabel:'🚀 Конвертировать',info:'Режим «Вся папка» обработает все PDF пакетно. Если PDF — это скан, текст не распознается.',ph:'Путь к файлу или папке…'},
 {id:'w2p',g:'Конвертация',ic:'🔄',t:'Word → PDF',sub:'Конвертация документов',mode:1,ftype:'docx',prog:1,act:'run_w2p',actLabel:'🚀 Конвертировать',info:'Кириллица поддерживается (шрифт DejaVu). Для идеального сохранения таблиц и картинок желательно иметь MS Word.',ph:'Путь к файлу или папке…'},
 {id:'img',g:'Конвертация',ic:'🖼',t:'Фото → PDF',sub:'Сборка из изображений',list:'multiimg',act:'run_img',actLabel:'🚀 Создать PDF',info:'JPG, PNG, BMP, GIF, TIFF. Порядок файлов = порядок страниц в PDF.'},
 {id:'split',g:'Инструменты',ic:'✂️',t:'Разделить PDF',sub:'На отдельные страницы',ftype:'pdf',act:'run_split',actLabel:'✂️ Разделить',info:'Каждая страница сохранится отдельным файлом в выбранную папку.'},
 {id:'merge',g:'Инструменты',ic:'🔗',t:'Объединить PDF',sub:'Склейка файлов',list:'multipdf',act:'run_merge',actLabel:'🔗 Объединить',info:'Файлы склеиваются в том порядке, в котором вы их добавили в список.'},
 {id:'comp',g:'Инструменты',ic:'📉',t:'Сжать PDF',sub:'Уменьшение размера',mode:1,ftype:'pdf',prog:1,act:'run_compress',actLabel:'📉 Сжать',info:'Сжатие удаляет невидимые дубликаты шрифтов и мусор. Качество текста не страдает.'},
 {id:'prot',g:'Инструменты',ic:'🛡',t:'Защита паролем',sub:'Шифрование 128 бит',ftype:'pdf',pass:1,act:'run_protect',actLabel:'🛡 Защитить',info:'Пароль потребуется для открытия файла. Не забудьте его!'},
 {id:'extr',g:'Страницы',ic:'📑',t:'Извлечь страницы',sub:'По диапазону',ftype:'pdf',range:1,act:'run_extract',actLabel:'📑 Извлечь',info:'Можно указывать сложные диапазоны: 1-3, 5, 8-10. Пустое поле = извлечь все.'},
 {id:'rot',g:'Страницы',ic:'📐',t:'Повернуть страницы',sub:'Ориентация',ftype:'pdf',range:1,angle:1,act:'run_rotate',actLabel:'📐 Повернуть',info:'Пустое поле страниц = повернуть абсолютно все страницы в документе.'},
 {id:'del',g:'Страницы',ic:'🗑',t:'Удалить страницы',sub:'Чистка документа',ftype:'pdf',range:1,act:'run_delete',actLabel:'🗑 Удалить',info:'Удалить все страницы нельзя — сработает защита от создания пустого файла.'},
 {id:'num',g:'Страницы',ic:'🔢',t:'Номера страниц',sub:'Нумерация',ftype:'pdf',act:'run_numbers',actLabel:'🔢 Добавить номера',info:'Номера ставятся внизу по центру каждой страницы. Шрифт Helvetica.'}];

let cur=null,modeVal=0,files=[];
let stats = JSON.parse(localStorage.getItem('pdf_stats')) || {ops:0, oks:0};
let recents = JSON.parse(localStorage.getItem('pdf_recents')) || [];

const $=id=>document.getElementById(id);

function escapeHTML(str) {
  return str.replace(/[&<>'"]/g, tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag]));
}

function updateStatsUI() {
  $('s1').textContent = stats.ops;
  $('s2').textContent = stats.oks;
}

function drawRecents() {
  const e = $('rec');
  if (recents.length === 0) { e.innerHTML = '<div>пока пусто</div>'; return; }
  e.innerHTML = recents.map(r => `<div onclick="pywebview.api.open_folder('${r.path.replace(/\\/g, '\\\\')}')" title="Открыть папку">${escapeHTML(r.name)}</div>`).join('');
}

function nav(){
  let h='',g='';
  S.forEach(s=>{
    if(s.g!==g){g=s.g;h+=`<div class="grp">${g.toUpperCase()}</div>`}
    h+=`<button data-id="${s.id}">${s.ic} ${s.t}</button>`
  });
  $('nav').innerHTML=h;
  $('nav').querySelectorAll('button').forEach(b=>b.onclick=()=>render(b.dataset.id));
}

function render(id){
  cur=S.find(s=>s.id===id);modeVal=0;files=[];
  $('ttl').textContent=cur.t;$('sub').textContent=cur.sub;$('tips').textContent=cur.info;
  $('nav').querySelectorAll('button').forEach(b=>b.classList.toggle('on',b.dataset.id===id));
  let h='';
  if(cur.mode)h+=`<div class="row"><span class="lbl">Режим:</span><div class="seg" id="seg"><button class="on" data-m="0">Один файл</button><button data-m="1">Вся папка</button></div></div>`;
  if(cur.list)h+=`<div class="listbox" id="lb">Список пуст — добавьте файлы</div><div class="row"><button class="btn yellow" id="add">＋ Добавить</button><button class="btn ghost" id="clr">Очистить</button></div>`;
  else h+=`<input id="path" placeholder="${cur.ph||'Путь к файлу…'}"><div class="row"><button class="btn dark" id="brw">📁 Обзор</button></div>`;
  
  // DRAG & DROP ЗОНА
  if(!cur.list) {
      h+=`<div class="dropzone" id="dropzone"><b>Перетащите файл сюда</b>или нажмите «Обзор»</div>`;
  }

  if(cur.range)h+=`<div><div class="lbl" style="margin-bottom:6px">Страницы (напр. 1-3, 7)</div><input id="rng" placeholder="пусто = все"></div>`;
  if(cur.angle)h+=`<div class="row"><span class="lbl">Угол:</span><select id="ang"><option value="90">90°</option><option value="180">180°</option><option value="270">270°</option></select></div>`;
  if(cur.pass)h+=`<div><div class="lbl" style="margin-bottom:6px">Пароль</div><input id="pwd" type="password" placeholder="Придумайте пароль"></div>`;
  h+=`<div class="row" style="margin-top:4px"><button class="btn primary" id="go">${cur.actLabel}</button></div>`;
  if(cur.prog)h+=`<div class="progress"><i id="bar"></i></div>`;
  h+=`<div class="status" id="st">Готов к работе</div>`;
  
  $('panel').innerHTML=h;
  wire();
  drawRecents();
}

function wire(){
  if($('seg'))$('seg').querySelectorAll('button').forEach(b=>b.onclick=()=>{
    modeVal=+b.dataset.m;
    $('seg').querySelectorAll('button').forEach(x=>x.classList.toggle('on',x===b));
  });
  if($('brw'))$('brw').onclick=async()=>{
    const k=(cur.mode&&modeVal===1)?'folder':cur.ftype;
    const p=await pywebview.api.browse(k);
    if(p){$('path').value=p;st('Файл выбран','ok');}
  };
  if($('add'))$('add').onclick=async()=>{
    const r=await pywebview.api.browse(cur.list);
    if(r){files=files.concat(r);drawFiles();}
  };
  if($('clr'))$('clr').onclick=()=>{files=[];drawFiles();};
  
  // ЛОГИКА DRAG & DROP
  const dz = $('dropzone');
  if(dz) {
    dz.ondragover = (e) => { e.preventDefault(); dz.classList.add('dragover'); };
    dz.ondragleave = () => dz.classList.remove('dragover');
    dz.ondrop = (e) => {
      e.preventDefault(); dz.classList.remove('dragover');
      // В PyWebView файлы из OS не всегда передаются в dataTransfer.
      // Поэтому вызываем диалог выбора для надежности.
      $('brw').click(); 
    };
    dz.onclick = () => $('brw').click();
  }

  $('go').onclick=go;
}

function drawFiles(){
  $('lb').innerHTML=files.length
    ? files.map(f=>escapeHTML(f.split(/[\\/]/).pop())).join('<br>')
    : 'Список пуст — добавьте файлы';
}

function st(t,k){const e=$('st');if(e){e.textContent=t;e.className='status '+(k||'');}}

async function go(){
  const b=$('go');b.disabled=true;b.classList.add('loading');st('Выполняется…');
  if($('bar'))$('bar').style.width='0%';
  const d={
    path:$('path')?$('path').value:'',
    mode:modeVal,files:files,
    range:$('rng')?$('rng').value:'',
    angle:$('ang')?$('ang').value:90,
    pwd:$('pwd')?$('pwd').value:''
  };
  try{
    const r=await pywebview.api[cur.act](d);
    stats.ops++; if(r.ok) stats.oks++;
    localStorage.setItem('pdf_stats', JSON.stringify(stats));
    updateStatsUI();

    if(r.ok){
      st('✓ '+r.msg,'ok');toast('✓ '+r.msg,false);
      if(r.file) {
        rec(r.file, r.is_dir);
      }
    }
    else if(!r.silent){st('✗ '+r.msg,'err');toast('✗ '+r.msg,true);}
    else st('Отменено');
  }catch(e){st('✗ '+e,'err');}
  b.disabled=false;b.classList.remove('loading');
}

function rec(path, is_dir){
  const name = is_dir ? path.split(/[\\/]/).pop() + ' (папка)' : path.split(/[\\/]/).pop();
  recents = recents.filter(r => r.path !== path); // Удаляем дубликаты
  recents.unshift({name: name, path: path});
  if(recents.length > 6) recents.pop();
  localStorage.setItem('pdf_recents', JSON.stringify(recents));
  drawRecents();
}

function toast(m,err){
  const t=document.createElement('div');
  t.className='toast'+(err?' err':'');
  t.textContent=m;
  document.body.appendChild(t);
  setTimeout(()=>t.remove(),4500);
}

window.UI={prog:(p,l)=>{
  const b=$('bar');if(b)b.style.width=(p*100)+'%';
  if(l)st(l);
}};

async function askUpdate(){
  const r=await pywebview.api.check_update();
  if(!r.ok){toast('Нет связи с сервером обновлений',true);return;}
  if(!r.new){toast('У вас последняя версия ✓',false);return;}
  if(confirm(`Доступна версия ${r.new}\n${r.notes||''}\n\nСкачать и установить сейчас?`)){
    st('Обновление…');
    const res=await pywebview.api.do_update(r.url);
    if(!res.ok)toast('Ошибка загрузки: '+res.msg,true);
  }
}

window.addEventListener('pywebviewready',async()=>{
  nav();render('p2w');updateStatsUI();
  pywebview.api.get_version().then(v=>{$('ver').textContent='v'+v;});
  $('upd').onclick=askUpdate;
  
  // Горячая клавиша Ctrl+O
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
      e.preventDefault();
      if($('brw')) $('brw').click();
    }
    if (e.key === 'Enter' && $('go') && !$('go').disabled) {
      $('go').click();
    }
  });

  const r=await pywebview.api.check_update();
  if(r.ok&&r.new)toast(`⬆ Доступна версия ${r.new} — нажмите «Проверить обновления»`,false);
});