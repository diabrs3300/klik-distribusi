/**
 * main.js - JavaScript utama
 */

// Auto-hide flash messages setelah 5 detik
document.addEventListener('DOMContentLoaded', () => {
  const alerts = document.querySelectorAll('.alert');
  alerts.forEach(alert => {
    setTimeout(() => {
      alert.style.transition = 'opacity 0.5s';
      alert.style.opacity = '0';
      setTimeout(() => alert.remove(), 500);
    }, 5000);
  });

  // Check if we should reset form after upload success
  const resetFormId = sessionStorage.getItem('reset_upload_form_id');
  if (resetFormId) {
      const formToReset = document.getElementById(resetFormId);
      if (formToReset) {
          formToReset.reset();
          // Trigger any change events needed (e.g., for file inputs)
          formToReset.querySelectorAll('input').forEach(input => {
              input.dispatchEvent(new Event('change', { bubbles: true }));
          });
          sessionStorage.removeItem('reset_upload_form_id');
          sessionStorage.removeItem('saved_col_map'); // Also clear mapping on success
          sessionStorage.removeItem('saved_file_ext_ekspor');
          sessionStorage.removeItem('saved_file_ext_impor');
      }
  }

  // Restore saved mapping values if any (for error persistence)
  const savedMapping = sessionStorage.getItem('saved_col_map');
  if (savedMapping) {
      try {
          const mapping = JSON.parse(savedMapping);
          Object.keys(mapping).forEach(name => {
              const inputs = document.querySelectorAll(`input[name="${name}"]`);
              inputs.forEach(inp => {
                  inp.value = mapping[name];
                  
                  // Restore data-last-type for ekspor/impor inputs to prevent Smart Reset from overwriting
                  if (inp.classList.contains('eksporColInput')) {
                      const lastExt = sessionStorage.getItem('saved_file_ext_ekspor');
                      if (lastExt) inp.setAttribute('data-last-type', lastExt);
                  }
                  if (inp.classList.contains('imporColInput')) {
                      const lastExt = sessionStorage.getItem('saved_file_ext_impor');
                      if (lastExt) inp.setAttribute('data-last-type', lastExt);
                  }
              });
          });
      } catch (e) {
          console.error('Error restoring mapping:', e);
      }
  }
});

/**
 * Global handler for file uploads with real-time progress polling.
 */
function handleUploadWithProgress(formId, btnId, textId, spinId, progressContainerId, progressBarId, progressTextId, progressPercentId) {
    const form = document.getElementById(formId);
    if (!form) return;

    form.addEventListener('submit', async function (e) {
        e.preventDefault();

        const btn = document.getElementById(btnId);
        const text = document.getElementById(textId);
        const spin = document.getElementById(spinId);

        const progressContainer = document.getElementById(progressContainerId);
        const progressBar = document.getElementById(progressBarId);
        const pText = document.getElementById(progressTextId);
        const pPercent = document.getElementById(progressPercentId);

        // Individual file size check (4.5MB limit)
        const fileInput = form.querySelector('input[type="file"]');
        if (fileInput && fileInput.files.length > 0) {
            const fileSizeMB = fileInput.files[0].size / 1024 / 1024;
            if (fileSizeMB > 4.5) {
                alert('🔴 UKURAN FILE TERLALU BESAR!\n\nMaksimal ukuran yang diizinkan server adalah 4.5 MB.\nUkuran file Anda: ' + fileSizeMB.toFixed(2) + ' MB.\n\nSilakan hapus baris/kolom kosong sebelum mengupload ulang.');
                return;
            }
        }

        // UI Reset
        if (btn) btn.disabled = true;
        if (text) text.classList.add('d-none');
        if (spin) spin.classList.remove('d-none');

        if (progressContainer) {
            progressContainer.classList.remove('d-none');
            if (progressBar) {
                progressBar.style.width = '0%';
                progressBar.setAttribute('aria-valuenow', 0);
            }
            if (pPercent) pPercent.textContent = '0%';
            if (pText) pText.textContent = 'Menyiapkan data...';
        }

        const taskId = 'task_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
        const formData = new FormData(this);
        formData.append('task_id', taskId);

        // Save current mapping to sessionStorage for persistence on failure
        const mapping = {};
        form.querySelectorAll('input[name^="col_map["]').forEach(inp => {
            mapping[inp.name] = inp.value;
        });
        sessionStorage.setItem('saved_col_map', JSON.stringify(mapping));

        // Also save extension trackers if they exist
        const extEkspor = form.querySelector('#file_ext_ekspor');
        const extImpor = form.querySelector('#file_ext_impor');
        if (extEkspor) sessionStorage.setItem('saved_file_ext_ekspor', extEkspor.value);
        if (extImpor) sessionStorage.setItem('saved_file_ext_impor', extImpor.value);

        const pollInterval = setInterval(async () => {
            try {
                const res = await fetch(`/dia-brs/upload-progress/${taskId}`);
                if (res.ok) {
                    const data = await res.json();
                    const pct = data.percent || 0;
                    if (progressBar) {
                        progressBar.style.width = pct + '%';
                        progressBar.setAttribute('aria-valuenow', pct);
                    }
                    if (pPercent) pPercent.textContent = pct + '%';
                    if (data.message && pText) pText.textContent = data.message;
                }
            } catch (err) {
                console.warn('Polling error:', err);
            }
        }, 1000);

        try {
            const response = await fetch(form.action || window.location.href, {
                method: 'POST',
                body: formData,
                headers: {
                    'X-Requested-With': 'XMLHttpRequest'
                }
            });

            clearInterval(pollInterval);

            if (response.ok) {
                // Signal form reset and reload (Success path stays the same)
                sessionStorage.setItem('reset_upload_form_id', formId);
                location.reload(); 
            } else {
                let errorMsg = 'Gagal mengupload file. Silakan cek file Anda dan coba lagi.';
                try {
                    const data = await response.json();
                    if (data && data.message) {
                        errorMsg = '🔴 GAGAL: ' + data.message;
                    }
                } catch (e) {
                    console.warn('Could not parse error JSON:', e);
                }
                
                alert(errorMsg);

                // RESTORE UI - No reload so the file stays selected!
                if (btn) btn.disabled = false;
                if (text) text.classList.remove('d-none');
                if (spin) spin.classList.add('d-none');
                if (progressContainer) progressContainer.classList.add('d-none');
            }
        } catch (err) {
            if (typeof pollInterval !== 'undefined') clearInterval(pollInterval);
            console.error('Upload error:', err);
            alert('Terjadi kesalahan koneksi saat mengupload.');
            
            // RESTORE UI even on connection error
            if (btn) btn.disabled = false;
            if (text) text.classList.remove('d-none');
            if (spin) spin.classList.add('d-none');
            if (progressContainer) progressContainer.classList.add('d-none');
        }
    });
}
