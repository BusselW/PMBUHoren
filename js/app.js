// ES6 module entry point for PMBU Hoorzitting Notulen React app

import { SharePointService } from './services/sharepoint-service.js';
import { DEFAULT_CASE_VALUES as DEFAULT_CASE } from './config/constants.js';
import { addMinutesToTime as addMin } from './helpers/date-utils.js';
import { exportToExcel } from './helpers/excel-utils.js';

// React alias and hooks
const H = React.createElement;
const { useState, useEffect } = React;

// Initialize SharePoint service
const sp = new SharePointService();

// Status indicator component
const Status = ({ s, m }) =>
    H(
        'div',
        { className: `status-indicator ${
            s === 'success'
                ? 'status-success'
                : s === 'failed'
                ? 'status-failed'
                : 'status-checking'
        }` },
        H('div', { className: 'status-dot' }),
        H('span', null, m)
    );

// Modal component
const Modal = ({ open, title, txt, close }) =>
    open
        ? H(
              'div',
              { className: 'modal-overlay', onClick: close },
              H(
                  'div',
                  { className: 'modal', onClick: e => e.stopPropagation() },
                  H('h2', null, title),
                  H('p', null, txt),
                  H(
                      'button',
                      { className: 'btn btn-primary', onClick: close },
                      'Sluiten'
                  )
              )
          )
        : null;

// Main App
const App = () => {
    const [cs, setCs] = useState([{ ...DEFAULT_CASE }]);
    const [st, sm] = useState('checking');
    const [msg, mmsg] = useState('Controleren...');
    const [ml, mlc] = useState(false);
    const [mt, mmc] = useState('');
    const [ld, ldc] = useState(false);

    // Test SharePoint connection
    useEffect(() => {
        (async () => {
            try {
                await sp.testConnection();
                sm('success');
                mmsg('Verbonden');
            } catch (e) {
                sm('failed');
                mmsg('Mislukt');
            }
        })();
    }, []);

    const upd = (i, c) => {
        const a = [...cs];
        a[i] = c;
        setCs(a);
    };

    const save = async i => {
        const c = cs[i];
        if (!c.zaaknummer || !c.hearingDate) {
            mlc(true);
            mmc('Zaaknr en datum verplicht');
            return;
        }
        ldc(true);
        try {
            const ex = await sp.getCaseByZaaknummer(c.zaaknummer);
            if (ex) {
                c.id = ex.Id;
                await sp.updateItem(ex.Id, c);
                mmc('Bijgewerkt');
            } else {
                const r = await sp.createItem(c);
                c.id = r.Id;
                mmc('Opgeslagen');
            }
            upd(i, c);
        } catch (e) {
            mmc('Fout: ' + e.message);
            console.error(e);
        } finally {
            ldc(false);
            mlc(true);
        }
    };

    const add = () => setCs([...cs, { ...DEFAULT_CASE }]);

    return H(
        'div',
        { className: 'container' },
        H('header', { className: 'header' },
            H('h1', null, 'PMBU Hoorzitting Notulen'),
            H(Status, { s: st, m: msg }),
            H('button', { className: 'btn btn-primary', onClick: add }, '➕'),
            H('button', { className: 'btn btn-success', onClick: () => exportToExcel(cs) }, '📤')
        ),
        H('main', null,
            cs.map((c, i) =>
                H('div', { key: i, className: 'case-card' },
                    H('div', { className: 'form-group' },
                        H('label', null, 'Zaaknummer'),
                        H('input', {
                            className: 'form-control',
                            value: c.zaaknummer,
                            onChange: e => upd(i, { ...c, zaaknummer: e.target.value })
                        })
                    ),
                    H('div', { className: 'form-group' },
                        H('label', null, 'Datum'),
                        H('input', {
                            type: 'date',
                            className: 'form-control',
                            value: c.hearingDate,
                            onChange: e => {
                                const v = e.target.value;
                                upd(i, { ...c, hearingDate: v, endTime: addMin(c.startTime, 4) });
                            }
                        })
                    ),
                    H('div', { className: 'form-group' },
                        H('label', null, 'Start'),
                        H('input', {
                            type: 'time',
                            className: 'form-control',
                            value: c.startTime,
                            onChange: e =>
                                upd(i, { ...c, startTime: e.target.value, endTime: addMin(e.target.value, 4) })
                        })
                    ),
                    H('div', { className: 'form-group' },
                        H('label', null, 'Eind'),
                        H('input', {
                            type: 'time',
                            className: 'form-control',
                            value: c.endTime,
                            disabled: true
                        })
                    ),
                    H('button', { className: 'btn btn-primary', onClick: () => save(i) }, 'Opslaan')
                )
            )
        ),
        H(Modal, { open: ml, title: 'Info', txt: mt, close: () => mlc(false) }),
        ld &&
            H('div', { className: 'loading-overlay' },
                H('div', { className: 'loading-content' },
                    H('div', { className: 'loading-spinner' }),
                    H('p', null, 'Bezig...')
                )
            )
    );
};

// Render the app
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(H(App));
