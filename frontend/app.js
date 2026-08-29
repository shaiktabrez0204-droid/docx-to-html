/* DOCX → HTML — Vanilla JS frontend
   Handles: file validation, drag/drop, conversion via existing POST /upload,
   pipeline states, success iframe, errors, and accessibility.
   No frameworks. No build step.
*/
(function(){
  const form = document.getElementById('uploadForm');
  const fileInput = document.getElementById('docxInput');
  const dropzone = document.getElementById('dropzone');
  const selected = document.getElementById('selected');
  const fileNameEl = document.getElementById('fileName');
  const fileSizeEl = document.getElementById('fileSize');
  const convertWrap = document.getElementById('convertWrap');
  const convertBtn = document.getElementById('convertBtn');
  const pipeline = document.getElementById('pipeline');
  const converting = document.getElementById('converting');
  const successPanel = document.getElementById('successPanel');
  const postResult = document.getElementById('postResult');
  const cardHead = document.getElementById('cardHead');
  const statusLive = document.getElementById('statusLive');
  const alertBox = document.getElementById('alertBox');
  const stepsList = document.getElementById('convSteps');
  if(!form || !fileInput || !dropzone) return;

  let selectedFile = null;
  let stepTimer = null;

  function formatBytes(b){
    if(b<1024) return b+' B';
    if(b<1024*1024) return (b/1024).toFixed(1)+' KB';
    return (b/1024/1024).toFixed(1)+' MB';
  }
  function setStatus(msg){ if(statusLive) statusLive.textContent = msg; }
  function setStepActive(active){
    const steps = pipeline.querySelectorAll('.pipe-step');
    steps.forEach(s=>{
      const n = parseInt(s.dataset.step,10);
      s.classList.remove('active','done');
      if(n<active) s.classList.add('done');
      else if(n===active) s.classList.add('active');
    });
  }
  function showAlert(msg, kind){
    if(!alertBox) return;
    alertBox.textContent = msg;
    alertBox.className = kind==='ok'?'ok':'err';
    alertBox.style.display = 'block';
    alertBox.setAttribute('role','alert');
  }
  function hideAlert(){ if(alertBox){ alertBox.style.display='none'; alertBox.textContent=''; alertBox.className=''; } }

  function setSelected(file){
    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = formatBytes(file.size);
    selected.classList.add('visible');
    dropzone.classList.add('has-file');
    convertWrap.classList.add('visible');
    hideAlert();
    setStatus('File selected: '+file.name+' ready to convert');
    // re-trigger animation
    selected.style.animation='none'; void selected.offsetWidth; selected.style.animation='';
    setStepActive(1);
  }
  function clearSelected(){
    selectedFile=null;
    selected.classList.remove('visible');
    dropzone.classList.remove('has-file');
    convertWrap.classList.remove('visible');
    fileInput.value='';
    setStatus('File removed');
  }

  function validateFile(file){
    if(!file) return 'No file';
    if(!file.name.toLowerCase().endsWith('.docx')) return 'Only .docx files are supported';
    if(file.size > 100*1024*1024) return 'File exceeds 100MB limit';
    return null;
  }

  // File input change
  fileInput.addEventListener('change', function(){
    const f = this.files && this.files[0];
    if(!f) return;
    const err = validateFile(f);
    if(err){ showAlert(err,'err'); dropzone.classList.add('drag-invalid'); setTimeout(()=>dropzone.classList.remove('drag-invalid'),1800); return; }
    hideAlert();
    setSelected(f);
  });

  // Drag & drop helpers
  let dragCounter=0;
  function isValidDrag(e){
    const items = e.dataTransfer && e.dataTransfer.items;
    if(items && items.length){
      for(let i=0;i<items.length;i++){
        const f = items[i].getAsFile && items[i].getAsFile();
        if(f && !f.name.toLowerCase().endsWith('.docx')) return false;
      }
    }
    return true;
  }
  dropzone.addEventListener('dragenter', function(e){
    e.preventDefault();
    dragCounter++;
    const valid = isValidDrag(e);
    dropzone.classList.add('drag-over');
    if(!valid) dropzone.classList.add('drag-invalid');
    else dropzone.classList.remove('drag-invalid');
  });
  dropzone.addEventListener('dragover', function(e){
    e.preventDefault();
    if(e.dataTransfer) e.dataTransfer.dropEffect='copy';
    const valid = isValidDrag(e);
    if(!valid) dropzone.classList.add('drag-invalid');
    else dropzone.classList.remove('drag-invalid');
  });
  dropzone.addEventListener('dragleave', function(e){
    e.preventDefault();
    dragCounter--;
    if(dragCounter<=0){ dragCounter=0; dropzone.classList.remove('drag-over','drag-invalid'); }
  });
  dropzone.addEventListener('drop', function(e){
    e.preventDefault();
    dragCounter=0;
    dropzone.classList.remove('drag-over','drag-invalid');
    const files = e.dataTransfer && e.dataTransfer.files;
    if(!files || !files.length) return;
    const f = files[0];
    const err = validateFile(f);
    if(err){
      dropzone.classList.add('drag-invalid');
      showAlert(err,'err');
      setStatus(err);
      setTimeout(()=>dropzone.classList.remove('drag-invalid'),1800);
      return;
    }
    try{
      const dt = new DataTransfer();
      dt.items.add(f);
      fileInput.files = dt.files;
    }catch(_){ /* fallback for older */ }
    setSelected(f);
  });

  const removeBtn = document.getElementById('removeFile');
  if(removeBtn) removeBtn.addEventListener('click', function(e){ e.preventDefault(); clearSelected(); });

  dropzone.addEventListener('keydown', function(e){
    if(e.key==='Enter' || e.key===' '){ e.preventDefault(); fileInput.click(); }
  });

  function showConverting(){
    cardHead.style.display='none';
    dropzone.style.display='none';
    selected.style.display='none';
    convertWrap.style.display='none';
    hideAlert();
    converting.classList.add('visible');
    setStepActive(2);
    setStatus('Converting your document, please wait');
    const lis = stepsList.querySelectorAll('li');
    let idx=0;
    if(stepTimer) clearInterval(stepTimer);
    lis.forEach(li=>li.classList.remove('active','done'));
    stepTimer = setInterval(()=>{
      if(idx>0) lis[idx-1].classList.remove('active'), lis[idx-1].classList.add('done');
      if(idx<lis.length){ lis[idx].classList.add('active'); idx++; }
      else { clearInterval(stepTimer); }
    }, 420);
  }
  function showSuccess(docId, dlName, previewUrl, downloadUrl){
    if(stepTimer) clearInterval(stepTimer);
    converting.classList.remove('visible');
    successPanel.classList.add('visible');
    const sName = document.getElementById('successFileName');
    const sDl = document.getElementById('successDlName');
    if(sName && selectedFile) sName.textContent = selectedFile.name;
    if(sDl) sDl.textContent = dlName;
    const openBtn = document.getElementById('openPreviewBtn');
    const dlBtn = document.getElementById('downloadBtn');
    if(openBtn) openBtn.href = previewUrl;
    if(dlBtn){ dlBtn.href = downloadUrl; dlBtn.setAttribute('download', dlName); }
    setStepActive(3);
    setStatus('Conversion complete');
    postResult.classList.remove('hidden');
    if(!postResult.querySelector('iframe.preview')){
      postResult.innerHTML = ''
        + '<div class="result"><div class="result-head"><div class="success-icon"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M20 6L9 17l-5-5"/></svg></div><div><div class="result-title">Conversion complete</div><div class="result-sub">'+escapeHtml(dlName)+' ready</div></div></div>'
        + '<div class="result-actions"><a class="btn small" href="'+previewUrl+'" download="'+escapeHtml(dlName)+'">Download HTML</a><a class="btn small ghost" href="/">Convert another</a></div></div>'
        + '<div class="result-frame-wrap"><div class="result-bar"><span class="badge">Converted</span><a class="btn small" href="'+downloadUrl+'" download="'+escapeHtml(dlName)+'">Download HTML</a><a class="btn small ghost" href="/">Convert another</a></div>'
        + '<iframe class="preview" src="'+previewUrl+'" title="Converted document preview" loading="lazy"></iframe></div>';
    }
    setTimeout(()=>{ try{ postResult.scrollIntoView({behavior:'smooth', block:'start'});}catch(e){ postResult.scrollIntoView(); } }, 180);
  }
  function showError(msg){
    if(stepTimer) clearInterval(stepTimer);
    converting.classList.remove('visible');
    cardHead.style.display='';
    dropzone.style.display='';
    if(selectedFile){ selected.style.display='flex'; convertWrap.style.display='block'; }
    showAlert(msg || 'Conversion failed. Please try again.', 'err');
    setStepActive(1);
    setStatus('Conversion failed: '+msg);
    if(convertBtn){ convertBtn.disabled=false; convertBtn.textContent='Convert to HTML'; }
  }
  function escapeHtml(s){
    return (s||'').replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; });
  }

  form.addEventListener('submit', function(e){
    const f = selectedFile || (fileInput.files && fileInput.files[0]);
    if(!f){
      e.preventDefault();
      showAlert('Please choose a .docx file','err');
      return;
    }
    const err = validateFile(f);
    if(err){ e.preventDefault(); showAlert(err,'err'); return; }
    if(window.fetch && window.FormData){
      e.preventDefault();
      hideAlert();
      showConverting();
      convertBtn.disabled=true;
      convertBtn.innerHTML='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="animation:spin 0.8s linear infinite"><path d="M21 12a9 9 0 1 1-6.2-8.6"/></svg> Converting...';
      if(!document.getElementById('spinKF')){
        const s=document.createElement('style'); s.id='spinKF'; s.textContent='@keyframes spin{to{transform:rotate(360deg)}}';
        document.head.appendChild(s);
      }
      const fd = new FormData();
      fd.append('docx', f, f.name);
      fetch('/upload', {method:'POST', body:fd, headers: {'Accept':'text/html, application/json'}})
        .then(async (resp)=>{
          const ct = resp.headers.get('Content-Type')||'';
          if(!resp.ok){
            let msg='Conversion failed';
            if(ct.includes('json')){
              try{ const j=await resp.json(); msg=j.error||j.message||msg; }catch(_){ }
            } else {
              const text = await resp.text();
              const m = text.match(/<div[^>]*class="[^"]*alert[^"]*"[^>]*>([^<]+)<\/div>/);
              msg = m ? m[1].trim() : msg;
              if(text.includes('Only .docx')) msg='Only .docx files are accepted.';
              else if(text.includes('not a valid DOCX')) msg='The uploaded file is not a valid DOCX.';
              else if(text.includes('too large')) msg='File exceeds 100MB limit.';
            }
            throw new Error(msg);
          }
          if(ct.includes('application/json')){
            const j = await resp.json();
            if(!j.doc_id) throw new Error('Conversion succeeded but no preview available');
            const docId=j.doc_id;
            const dlName=j.download_name|| (f.name.replace(/\.docx$/i,'')+'.html');
            showSuccess(docId, dlName, '/preview/'+docId, '/download/'+docId);
            convertBtn.disabled=false; convertBtn.textContent='Convert to HTML';
            return;
          }
          const text = await resp.text();
          const idMatch = text.match(/\/preview\/([a-f0-9]+)/);
          const dlMatch = text.match(/download="([^"]+\.html)"/);
          const docId = idMatch ? idMatch[1] : null;
          const dlName = dlMatch ? dlMatch[1] : (f.name.replace(/\.docx$/i,'')+'.html');
          if(!docId) throw new Error('Conversion succeeded but no preview available');
          const previewUrl = '/preview/'+docId;
          const downloadUrl = '/download/'+docId;
          showSuccess(docId, dlName, previewUrl, downloadUrl);
          convertBtn.disabled=false;
          convertBtn.textContent='Convert to HTML';
        })
        .catch(err=>{
          showError(err.message || 'Conversion failed');
          convertBtn.disabled=false;
          convertBtn.textContent='Convert to HTML';
        });
    }
  });

  // Initial pipeline state
  setStepActive(1);

  // Global Escape handling for any overlay (future)
  document.addEventListener('keydown', function(e){
    if(e.key==='Escape'){
      // clear search if any, else close drawer if viewer (handled by viewer script)
    }
  });

  // Expose for tests
  window.__docxUI = { setSelected, clearSelected, validateFile, showConverting, showSuccess, showError, setStepActive };
})();
