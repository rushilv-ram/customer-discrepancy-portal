document.addEventListener('DOMContentLoaded', ()=>{
  const map = window.WAREHOUSE_MAP || {};

  // Bind warehouse lookup for a given invoice-item element
  function bindInvoiceItem(item){
    const invoiceInput = item.querySelector('input[name="invoice_no"]');
    const whField = item.querySelector('input[name="warehouse_auto"]');
    const whCodeField = item.querySelector('input[name="warehouse_code_auto"]');
    const fileInput = item.querySelector('input[name="invoice_image"]');
    const previewImg = item.querySelector('.invoice-preview');
    const meta = item.querySelector('.file-meta');

    if(invoiceInput){
      invoiceInput.addEventListener('input', ()=>{
        const v = (invoiceInput.value||'').trim();
        const k = v.slice(0,4);
        if(map[k]){
          whField.value = map[k][0];
          whCodeField.value = map[k][1];
        } else {
          whField.value = 'Unknown';
          whCodeField.value = 'Unknown';
        }
      });
    }

    if(fileInput){
      fileInput.addEventListener('change', ()=>{
        const f = fileInput.files && fileInput.files[0];
        if(!f){ previewImg.src=''; meta.textContent = 'No file selected'; return }
        meta.textContent = `${f.name} · ${(f.size/1024).toFixed(0)} KB`;
        const reader = new FileReader();
        reader.onload = (e)=>{ previewImg.src = e.target.result };
        reader.readAsDataURL(f);
      });
    }

    const removeBtn = item.querySelector('.remove-invoice');
    if(removeBtn){
      removeBtn.addEventListener('click', ()=>{
        const list = document.getElementById('invoice-list');
        if(list.querySelectorAll('.invoice-item').length <= 1){
          // don't remove last item; clear fields instead
          item.querySelectorAll('input').forEach(i=>{ if(i.type!=='file') i.value = ''; });
          const img = item.querySelector('.invoice-preview'); if(img) img.src = '';
          const meta = item.querySelector('.file-meta'); if(meta) meta.textContent = 'No file selected';
          return;
        }
        item.remove();
      });
    }
  }

  // Initialize existing invoice items
  document.querySelectorAll('.invoice-item').forEach(bindInvoiceItem);

  // Add invoice button
  const addBtn = document.getElementById('add-invoice');
  if(addBtn){
    addBtn.addEventListener('click', ()=>{
      const tmpl = document.getElementById('invoice-template');
      const clone = tmpl.content.cloneNode(true);
      const list = document.getElementById('invoice-list');
      list.appendChild(clone);
      // Bind the newly added item (last one)
      const items = list.querySelectorAll('.invoice-item');
      const newItem = items[items.length-1];
      bindInvoiceItem(newItem);
    });
  }

  // Global previews for LR and Discrepancy (existing ids)
  function setupPreviewById(fileInputName, imgId, metaId){
    const input = document.querySelector(`input[name="${fileInputName}"]`);
    const img = document.getElementById(imgId);
    const meta = document.getElementById(metaId);
    if(!input) return;
    input.addEventListener('change', ()=>{
      const f = input.files && input.files[0];
      if(!f){ img.src=''; meta.textContent = 'No file selected'; return }
      meta.textContent = `${f.name} · ${(f.size/1024).toFixed(0)} KB`;
      const reader = new FileReader();
      reader.onload = (e)=>{ img.src = e.target.result };
      reader.readAsDataURL(f);
    });
  }

  setupPreviewById('lr_image','lr_preview','lr_meta');
  setupPreviewById('discrepancy_image','discrepancy_preview','discrepancy_meta');
});