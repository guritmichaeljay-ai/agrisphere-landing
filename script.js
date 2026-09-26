document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) window.lucide.createIcons();

  const header = document.querySelector('.site-header');
  const handleScroll = () => {
    header.classList.toggle('is-scrolled', window.scrollY > 12);
  };

  handleScroll();
  window.addEventListener('scroll', handleScroll, { passive: true });

  const form = document.querySelector('#access-form');
  const exportButton = document.querySelector('#export-records');
  const formMessage = document.querySelector('#form-message');
  const storageKey = 'agrisphere-access-records';
  const today = () => {
    const date = new Date();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${date.getFullYear()}-${month}-${day}`;
  };
  const googleSheetsEndpoint = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
  const isGoogleSheetsConfigured = googleSheetsEndpoint && !googleSheetsEndpoint.startsWith('PASTE_');
  let records = [];
  try {
    const savedRecords = JSON.parse(localStorage.getItem(storageKey) || '[]');
    records = Array.isArray(savedRecords) ? savedRecords : [];
  } catch {
    localStorage.removeItem(storageKey);
  }
  const escapeHtml = (value) => String(value).replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' })[character]);
  const exportRecords = () => {
    if (!records.length || !window.XLSX) return;
    const rows = records.map((record) => ({
      Name: record.name,
      Email: record.email,
      'Access type': record.accessType,
      'Farm / organization': record.organization,
      Purpose: record.purpose,
      Date: record.date,
      Status: 'Recorded'
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 22 }, { wch: 30 }, { wch: 20 }, { wch: 26 },
      { wch: 22 }, { wch: 14 }, { wch: 14 }
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Access Records');
    XLSX.writeFile(workbook, `agrisphere-access-records-${today()}.xlsx`);
  };

  const renderRecords = () => {
    exportButton.disabled = records.length === 0;
  };

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const data = new FormData(form);
    const record = {
      name: data.get('name'), email: data.get('email'), accessType: data.get('accessType'),
      organization: data.get('organization'), purpose: data.get('purpose'), date: today()
    };
    const duplicateRegistration = record.accessType === 'Registered person'
      && records.some((savedRecord) => savedRecord.accessType === 'Registered person'
        && String(savedRecord.email || '').toLowerCase() === record.email.toLowerCase());
    if (duplicateRegistration) {
      formMessage.textContent = 'This person is already registered. No new file was created.';
      return;
    }
    records.unshift(record);
    try {
      localStorage.setItem(storageKey, JSON.stringify(records));
    } catch {
      formMessage.textContent = 'This browser could not save the record locally.';
      return;
    }
    renderRecords();
    form.reset();
    formMessage.textContent = `${records.length} record${records.length === 1 ? '' : 's'} saved. Use Download Excel file for the complete workbook.`;
    if (isGoogleSheetsConfigured) {
      fetch(googleSheetsEndpoint, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(record) })
        .then(() => { formMessage.textContent = 'Record added to the Google Sheets collection.'; })
        .catch(() => { formMessage.textContent = 'Saved locally, but Google Sheets sync failed.'; });
    } else {
      formMessage.textContent = `${records.length} record${records.length === 1 ? '' : 's'} saved locally. Replace the Apps Script URL in script.js to use one shared file.`;
    }
    setTimeout(() => { formMessage.textContent = ''; }, 4000);
  });

  exportButton.addEventListener('click', exportRecords);

  renderRecords();
});
