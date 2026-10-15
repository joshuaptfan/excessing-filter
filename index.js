var sites = {};
var siteData = {};
var siteCode;
const types = [
  '',
  'Laptop',
  'UPS',
  'Server',
  'Other IT Hardware',
  'Monitor',
  'Switch',
  'Desktop',
  'Printer/Fax/Scanner',
  'Phones',
  'Docking Station',
  'TV',
  'Keyboards',
  'Cables'
];

document.onreadystatechange = function () {
  document.addEventListener('dragover', dragOverFile);
  document.addEventListener('drop', dropFile);
  document.querySelector('header > input').addEventListener('change', e => readFile(e.target.files[0]));
  document.querySelector('search > select').addEventListener('change', e => loadSite(e.target.value, true));
  document.getElementById('save-btn').addEventListener('click', saveChanges);
  document.getElementById('revert-btn').addEventListener('click', revertChanges);
  document.getElementById('copy-btn').addEventListener('click', copyCounts);
  document.getElementById('equipment').addEventListener('change', e => setCount(e.target.name, parseInt(e.target.value)));
  document.getElementById('other-IT').addEventListener('click', e => e.target.type === 'button' && addToType(e));
}

function dragOverFile(e) {
  e.stopPropagation();
  if (e.dataTransfer.items.length > 0) {
    e.preventDefault();
    if (e.dataTransfer.items[0].kind === 'file' && e.dataTransfer.items[0].type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      e.dataTransfer.dropEffect = 'copy';
      const dropTarget = document.getElementById('drop-target');
      dropTarget.style.display = 'block';
      dropTarget.addEventListener('dragleave', () => {
        dropTarget.style.display = 'none';
      }, { once: true });
    } else
      e.dataTransfer.dropEffect = 'none';
  }
}

function dropFile(e) {
  e.stopPropagation();
  e.preventDefault();
	document.getElementById('drop-target').style.display = 'none';
  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(e.dataTransfer.items[0].getAsFile());
  document.querySelector('header > input').files = dataTransfer.files;
  readFile(dataTransfer.files[0]);
}

function readFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const workbook = XLSX.read(reader.result, { type: 'array' });
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    const siteCodeFrag = document.createDocumentFragment();
    let unsortedSites = {};
    for (let i = 0; i < data.length; i++) {
      const code = data[i].SiteCode;
      const deviceType = data[i]['Device Type'];
      if (!Object.hasOwn(unsortedSites, code))
        unsortedSites[code] = { ['Equipment Count']: 0 };
      unsortedSites[code]['Equipment Count']++;
      if (deviceType !== 'Other IT Hardware')
        unsortedSites[code][deviceType] = (unsortedSites[code][deviceType] || 0) + 1;
      else {
        unsortedSites[code]['Other IT'] ??= [];
        unsortedSites[code]['Other IT'].push({
          serial: data[i]['Serial Number'],
          count: data[i]['Count'],
          desc: data[i]['Description']
        });
      }
    }

    // Sort site codes
    sites = Object.keys(unsortedSites)
      .sort()
      .reduce((obj, key) => {
        obj[key] = unsortedSites[key];
        return obj;
      }, {});

    // Populate site code list
    for (const code of Object.keys(sites)) {
      const option = document.createElement('option');
      option.innerHTML = code;
      siteCodeFrag.appendChild(option);
    }
    document.querySelector('search > select').appendChild(siteCodeFrag);
  };
  reader.readAsArrayBuffer(file);
}

function loadSite(code, isRevert) {
  const equipment = document.querySelector('#equipment > tbody');
  const otherIT = document.querySelector('#other-IT > tbody');
  const equipmentFrag = document.createDocumentFragment();
  const otherITFrag = document.createDocumentFragment();
  siteCode = code;
  document.querySelector('main > h2').innerHTML = 'Site ' + code;
  // Clear existing elements
  while (equipment.childNodes.length)
    equipment.removeChild(equipment.lastChild);
  while (otherIT.childNodes.length)
    otherIT.removeChild(otherIT.lastChild);

  if (isRevert)
    siteData = structuredClone(sites[code]);
  for (let [deviceType, value] of Object.entries(siteData)) {
    if (!Array.isArray(value)) {
      // Populate equipment list
      const tr = document.createElement('tr');
      const nameTH = document.createElement('th');
      const numInput = document.createElement('input');
      const countTD = document.createElement('td');
      nameTH.innerHTML = deviceType;
      if (deviceType !== 'Equipment Count') {
        numInput.type = 'number';
        numInput.name = deviceType;
        numInput.value = value;
        countTD.appendChild(numInput);
      } else
        countTD.innerHTML = value;
      tr.append(nameTH, countTD);
      equipmentFrag.appendChild(tr);
    } else {
      // Populate other IT hardware list
      for (let i = 0; i < value.length; i++) {
        const tr = document.createElement('tr');
        const select = document.createElement('select');
        const applyBtn = document.createElement('button');
        const typeTD = document.createElement('td');
        const serialTD = document.createElement('td');
        const numInput = document.createElement('input');
        const countTD = document.createElement('td');
        const descTD = document.createElement('td');
        tr.name = i;
        for (const type of types) {
          const option = document.createElement('option');
          option.innerHTML = type;
          select.appendChild(option);
        }
        applyBtn.type = 'button';
        applyBtn.innerHTML = '+';
        typeTD.append(select, applyBtn);
        serialTD.innerHTML = value[i].serial;
        numInput.type = 'number';
        numInput.value = value[i].count;
        descTD.innerHTML = value[i].desc;
        countTD.appendChild(numInput);
        tr.append(typeTD, serialTD, countTD, descTD);
        otherITFrag.appendChild(tr);
      }
    }
  }
  equipment.appendChild(equipmentFrag);
  otherIT.appendChild(otherITFrag);
}

function saveChanges() {
  sites[siteCode] = structuredClone(siteData);
}

function revertChanges() {
  loadSite(siteCode, true);
}

async function copyCounts() {
  let text = '';
  for (let i = 1; i < types.length; i++)
    text += (siteData[types[i]] || 0) + '\t';
  console.info('Site ' + siteCode + '\t' + text);
  try {
    await navigator.clipboard.writeText(text);
  } catch (error) {
    console.error(error.message);
  }
}

function updateEquipmentCount() {
  let count = 0;
  for (const [k, v] of Object.entries(siteData))
    if (k !== 'Equipment Count' && !Array.isArray(v))
      count += v;
  count += siteData['Other IT'] ? siteData['Other IT'].length : 0;
  siteData['Equipment Count'] = count;
  document.querySelector('#equipment > tbody > tr > td').innerHTML = count;
}

function setCount(name, count) {
  siteData[name] = count;
  updateEquipmentCount();
}

function addToType(e) {
  const tr = e.target.parentNode.parentNode;
  const deviceType = tr.querySelector('select').value;
  const count = parseInt(tr.querySelector('input[type="number"]').value);
  if (!deviceType) return;
  siteData[deviceType] = (siteData[deviceType] || 0) + count;
  siteData['Other IT'].splice(parseInt(tr.name), 1)
  updateEquipmentCount();
  loadSite(siteCode);
}
