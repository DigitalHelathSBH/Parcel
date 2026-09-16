/**
 * ดรอปดาวน์ที่ค้นหาได้ (searchable combobox) แบบ vanilla JS ไม่ต้องพึ่ง library ภายนอก
 * ใช้กับ dropdown ที่มีรายการเยอะ เช่น ฝ่าย/แผนก/งาน ที่มีเป็นร้อยรายการ
 *
 * markup ที่ต้องมี (ดูตัวอย่างการใช้งานใน index.html / depreciation.html):
 *   <div id="{id}_wrap" class="relative">
 *     <input type="hidden" name="..." id="{id}_value">
 *     <button type="button" id="{id}_button">...</button>
 *     <div id="{id}_panel" class="hidden ...">
 *       <input type="text" id="{id}_search">
 *       <ul id="{id}_list"></ul>
 *     </div>
 *   </div>
 */
function makeSearchableSelect(id, { placeholder, onChange } = {}) {
  const wrap = document.getElementById(`${id}_wrap`);
  const valueInput = document.getElementById(`${id}_value`);
  const button = document.getElementById(`${id}_button`);
  const buttonLabel = button.querySelector('span');
  const panel = document.getElementById(`${id}_panel`);
  const search = document.getElementById(`${id}_search`);
  const list = document.getElementById(`${id}_list`);

  let items = [];

  function renderList(filterText = '') {
    list.innerHTML = '';
    const f = filterText.trim().toLowerCase();
    const filtered = f
      ? items.filter(it => it.label.toLowerCase().includes(f) || it.value.toLowerCase().includes(f))
      : items;

    const allLi = document.createElement('li');
    allLi.className = 'px-3 py-2 cursor-pointer hover:bg-brand-50 text-zinc-500';
    allLi.textContent = placeholder;
    allLi.addEventListener('click', () => selectItem('', placeholder));
    list.appendChild(allLi);

    if (filtered.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'px-3 py-2 text-zinc-400';
      empty.textContent = 'ไม่พบข้อมูล';
      list.appendChild(empty);
    }

    for (const it of filtered) {
      const li = document.createElement('li');
      li.className = 'px-3 py-2 cursor-pointer hover:bg-brand-50';
      li.textContent = it.label;
      li.addEventListener('click', () => selectItem(it.value, it.label));
      list.appendChild(li);
    }
  }

  function selectItem(value, label) {
    valueInput.value = value;
    buttonLabel.textContent = value ? label : placeholder;
    buttonLabel.classList.toggle('text-zinc-400', !value);
    closePanel();
    if (onChange) onChange(value);
  }

  function openPanel() {
    if (button.disabled) return;
    panel.classList.remove('hidden');
    search.value = '';
    renderList();
    search.focus();
  }

  function closePanel() {
    panel.classList.add('hidden');
  }

  button.addEventListener('click', () => {
    panel.classList.contains('hidden') ? openPanel() : closePanel();
  });
  search.addEventListener('input', () => renderList(search.value));
  document.addEventListener('click', (e) => {
    if (!wrap.contains(e.target)) closePanel();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closePanel();
  });

  return {
    setItems(newItems) {
      items = newItems;
    },
    reset() {
      valueInput.value = '';
      buttonLabel.textContent = placeholder;
      buttonLabel.classList.add('text-zinc-400');
      items = [];
    },
    setDisabled(disabled) {
      button.disabled = disabled;
      button.classList.toggle('opacity-50', disabled);
      button.classList.toggle('cursor-not-allowed', disabled);
      if (disabled) closePanel();
    },
    getValue() {
      return valueInput.value;
    },
  };
}
