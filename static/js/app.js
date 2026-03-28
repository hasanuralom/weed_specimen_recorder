/* ═══════════════════════════════════════════════════
   Weed Specimen Recorder — Application Logic
   ═══════════════════════════════════════════════════ */

const App = {
    specimens: [],
    fileInput: null,

    init() {
        this.fileInput = document.getElementById('fileInput');
        this.setupUploadZone();
        this.setupEventListeners();
        this.loadProject();
    },

    /* ── Upload Zone ── */
    setupUploadZone() {
        const zone = document.getElementById('uploadZone');
        const events = ['dragenter', 'dragover', 'dragleave', 'drop'];

        events.forEach(e => {
            zone.addEventListener(e, (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
            });
        });

        ['dragenter', 'dragover'].forEach(e => {
            zone.addEventListener(e, () => zone.classList.add('drag-over'));
        });

        ['dragleave', 'drop'].forEach(e => {
            zone.addEventListener(e, () => zone.classList.remove('drag-over'));
        });

        zone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length) this.uploadFiles(files);
        });

        zone.addEventListener('click', () => this.fileInput.click());
        this.fileInput.addEventListener('change', () => {
            if (this.fileInput.files.length) {
                this.uploadFiles(this.fileInput.files);
                this.fileInput.value = '';
            }
        });
    },

    /* ── Event Listeners ── */
    setupEventListeners() {
        document.getElementById('btnGenerate').addEventListener('click', () => this.generateDocx());
        document.getElementById('btnClear').addEventListener('click', () => this.clearAll());
        document.getElementById('btnSave').addEventListener('click', () => this.saveProject());
        document.getElementById('btnLoad').addEventListener('click', () => this.loadProject());
    },

    /* ── Upload Files ── */
    async uploadFiles(files) {
        const formData = new FormData();
        let count = 0;
        for (const f of files) {
            if (f.type.startsWith('image/')) {
                formData.append('images', f);
                count++;
            }
        }

        if (!count) {
            this.showToast('No valid image files selected.', 'error');
            return;
        }

        this.showProgress(`Uploading ${count} image${count > 1 ? 's' : ''}...`);

        try {
            const resp = await fetch('/upload', { method: 'POST', body: formData });
            const data = await resp.json();

            if (data.uploaded) {
                const startNum = this.specimens.length + 1;
                data.uploaded.forEach((img, i) => {
                    this.specimens.push({
                        id: crypto.randomUUID ? crypto.randomUUID() : Date.now().toString() + Math.random().toString(36),
                        filename: img.filename,
                        thumb: img.thumb,
                        original_name: img.original_name,
                        label: `Specimen ${startNum + i}`,
                        common_name: '',
                        scientific_name: '',
                        family: '',
                        type: '',
                        notes: ''
                    });
                });

                this.renderSpecimens();
                this.showToast(`${data.uploaded.length} image${data.uploaded.length > 1 ? 's' : ''} uploaded successfully!`, 'success');
            }
        } catch (err) {
            this.showToast('Upload failed: ' + err.message, 'error');
        }

        this.hideProgress();
    },

    /* ── Render Specimens ── */
    renderSpecimens() {
        const grid = document.getElementById('specimensGrid');
        const emptyState = document.getElementById('emptyState');

        if (!this.specimens.length) {
            grid.innerHTML = '';
            emptyState.style.display = 'block';
            this.updateStats();
            return;
        }

        emptyState.style.display = 'none';

        grid.innerHTML = this.specimens.map((spec, idx) => `
            <div class="specimen-card" data-id="${spec.id}" style="animation-delay: ${idx * 0.05}s">
                <div class="card-header">
                    <div>
                        <span class="specimen-label">${spec.label}</span>
                    </div>
                    <div style="display:flex;gap:6px;align-items:center;">
                        <span class="specimen-number">#${idx + 1}</span>
                        <button class="card-remove" onclick="App.removeSpecimen('${spec.id}')" title="Remove">✕</button>
                    </div>
                </div>
                <div class="card-image-wrapper">
                    <img src="/uploads/${spec.thumb}" alt="${spec.label}" loading="lazy" />
                </div>
                <div class="card-details">
                    <div class="form-group">
                        <label class="form-label">Specimen Label</label>
                        <input class="form-input" type="text" value="${this.escapeHtml(spec.label)}"
                            onchange="App.updateField('${spec.id}', 'label', this.value)" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Common Name</label>
                        <input class="form-input" type="text" placeholder="e.g. Wild Oat"
                            value="${this.escapeHtml(spec.common_name)}"
                            onchange="App.updateField('${spec.id}', 'common_name', this.value)" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Scientific Name</label>
                        <input class="form-input scientific" type="text" placeholder="e.g. Avena fatua"
                            value="${this.escapeHtml(spec.scientific_name)}"
                            onchange="App.updateField('${spec.id}', 'scientific_name', this.value)" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Family</label>
                        <input class="form-input" type="text" placeholder="e.g. Poaceae"
                            value="${this.escapeHtml(spec.family)}"
                            onchange="App.updateField('${spec.id}', 'family', this.value)" />
                    </div>
                    <div class="form-group">
                        <label class="form-label">Type</label>
                        <select class="form-select" onchange="App.updateField('${spec.id}', 'type', this.value)">
                            <option value="" ${spec.type === '' ? 'selected' : ''}>— Select Type —</option>
                            <option value="Grass" ${spec.type === 'Grass' ? 'selected' : ''}>🌾 Grass</option>
                            <option value="Broadleaf" ${spec.type === 'Broadleaf' ? 'selected' : ''}>🍃 Broadleaf</option>
                            <option value="Sedge" ${spec.type === 'Sedge' ? 'selected' : ''}>🌿 Sedge</option>
                            <option value="Monocot" ${spec.type === 'Monocot' ? 'selected' : ''}>🌱 Monocot</option>
                            <option value="Dicot" ${spec.type === 'Dicot' ? 'selected' : ''}>🌸 Dicot</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label class="form-label">Notes (optional)</label>
                        <textarea class="form-textarea" placeholder="Additional notes..."
                            onchange="App.updateField('${spec.id}', 'notes', this.value)">${this.escapeHtml(spec.notes)}</textarea>
                    </div>
                </div>
            </div>
        `).join('');

        this.updateStats();
    },

    /* ── Update Field ── */
    updateField(id, field, value) {
        const spec = this.specimens.find(s => s.id === id);
        if (spec) spec[field] = value;
    },

    /* ── Remove Specimen ── */
    removeSpecimen(id) {
        const idx = this.specimens.findIndex(s => s.id === id);
        if (idx === -1) return;

        const card = document.querySelector(`[data-id="${id}"]`);
        if (card) {
            card.style.transition = 'all 0.3s ease';
            card.style.opacity = '0';
            card.style.transform = 'scale(0.9)';
        }

        setTimeout(() => {
            this.specimens.splice(idx, 1);
            // Re-label
            this.specimens.forEach((s, i) => {
                if (s.label.startsWith('Specimen ')) {
                    s.label = `Specimen ${i + 1}`;
                }
            });
            this.renderSpecimens();
        }, 300);
    },

    /* ── Generate DOCX ── */
    async generateDocx() {
        if (!this.specimens.length) {
            this.showToast('No specimens to generate. Upload images first!', 'error');
            return;
        }

        const title = document.getElementById('docTitle').value || 'Weed Specimen Record';
        const cols = parseInt(document.getElementById('colsPerRow').value) || 3;

        this.showProgress('Generating Word document...');

        try {
            const resp = await fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    title: title,
                    cols_per_row: cols,
                    specimens: this.specimens.map(s => ({
                        filename: s.filename,
                        label: s.label,
                        common_name: s.common_name,
                        scientific_name: s.scientific_name,
                        family: s.family,
                        type: s.type,
                        notes: s.notes
                    }))
                })
            });

            const data = await resp.json();

            if (data.filename) {
                this.hideProgress();
                this.showToast('Word document generated! Downloading...', 'success');
                // Trigger download
                const a = document.createElement('a');
                a.href = `/download/${data.filename}`;
                a.download = 'Weed_Specimen_Record.docx';
                document.body.appendChild(a);
                a.click();
                a.remove();
            } else {
                throw new Error(data.error || 'Generation failed');
            }
        } catch (err) {
            this.showToast('Generation failed: ' + err.message, 'error');
        }

        this.hideProgress();
    },

    /* ── Save / Load Project ── */
    async saveProject() {
        if (!this.specimens.length) {
            this.showToast('Nothing to save.', 'info');
            return;
        }

        try {
            await fetch('/save-project', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    specimens: this.specimens,
                    title: document.getElementById('docTitle').value
                })
            });
            this.showToast('Project saved locally!', 'success');
        } catch (err) {
            this.showToast('Save failed: ' + err.message, 'error');
        }
    },

    async loadProject() {
        try {
            const resp = await fetch('/load-project');
            if (resp.ok) {
                const data = await resp.json();
                if (data.specimens && data.specimens.length) {
                    this.specimens = data.specimens;
                    if (data.title) {
                        document.getElementById('docTitle').value = data.title;
                    }
                    this.renderSpecimens();
                    this.showToast(`Loaded ${data.specimens.length} saved specimens.`, 'info');
                }
            }
        } catch (err) {
            // Silently fail on load — no saved project
        }
    },

    /* ── Clear All ── */
    async clearAll() {
        if (!this.specimens.length) return;

        if (!confirm('Are you sure you want to remove all specimens? This cannot be undone.')) return;

        try {
            await fetch('/clear', { method: 'POST' });
        } catch (err) { /* ignore */ }

        this.specimens = [];
        this.renderSpecimens();
        this.showToast('All specimens cleared.', 'info');
    },

    /* ── Stats ── */
    updateStats() {
        document.getElementById('statTotal').textContent = this.specimens.length;
        document.getElementById('statFilled').textContent =
            this.specimens.filter(s => s.common_name || s.scientific_name).length;
        document.getElementById('statTypes').textContent =
            new Set(this.specimens.filter(s => s.type).map(s => s.type)).size;
    },

    /* ── Toast ── */
    showToast(message, type = 'info') {
        const container = document.getElementById('toastContainer');
        const icons = { success: '✓', error: '✕', info: 'ℹ' };
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `<span style="font-size:18px">${icons[type] || 'ℹ'}</span> ${message}`;
        container.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100px)';
            toast.style.transition = 'all 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    },

    /* ── Progress ── */
    showProgress(text = 'Processing...') {
        const overlay = document.getElementById('progressOverlay');
        document.getElementById('progressText').textContent = text;
        overlay.classList.add('active');
    },

    hideProgress() {
        document.getElementById('progressOverlay').classList.remove('active');
    },

    /* ── Helpers ── */
    escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }
};

// ── Initialize on DOM Ready ──
document.addEventListener('DOMContentLoaded', () => App.init());
