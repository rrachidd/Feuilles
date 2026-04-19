import React, { useEffect, useState } from 'react';
import { auth, loginWithGoogle, logout, db } from './lib/firebase';
import { onAuthStateChanged, User } from 'firebase/auth';
import { collection, getDocs, setDoc, doc, writeBatch, deleteDoc, query, where, getDoc } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { LogOut, LogIn, Settings } from 'lucide-react';

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  // App States
  const [allStudents, setAllStudents] = useState<any[]>([]);
  const [filteredStudents, setFilteredStudents] = useState<any[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);
  const [colCount, setColCount] = useState(0);
  const [sampleRows, setSampleRows] = useState<any[]>([]);
  const [fileName, setFileName] = useState('');
  
  // UI States
  const [activeTab, setActiveTab] = useState<'students' | 'absence'>('students');
  const [mappingSuccess, setMappingSuccess] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [isSavingSettings, setIsSavingSettings] = useState(false);
  const [stats, setStats] = useState({ tot: 0, sh: 0, mal: 0, fem: 0 });
  const [search, setSearch] = useState('');
  const [filterSheet, setFilterSheet] = useState('');
  const [filterGender, setFilterGender] = useState('');

  const [colMap, setColMap] = useState({ code: 1, lastname: 2, firstname: 3, gender: 4, dept: 5 });

  // Absence Settings
  const [aDir, setaDir] = useState('');
  const [aSch, setaSch] = useState('');
  const [aTri, setaTri] = useState('الثلاثي الأول');
  const [aFrom, setaFrom] = useState('');
  const [aTo, setaTo] = useState('');
  const [aTeach, setaTeach] = useState('');
  const [aHalf, setaHalf] = useState(0);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [absenceHtml, setAbsenceHtml] = useState<string>('');

  useEffect(() => {
    const today = new Date();
    const d = today.getDay();
    const sat = new Date(today);
    sat.setDate(today.getDate() - ((d + 1) % 7));
    const mon = new Date(sat);
    mon.setDate(sat.getDate() + 2);
    setaFrom(sat.toISOString().split('T')[0]);
    setaTo(mon.toISOString().split('T')[0]);

    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      setUser(user);
      if (user) {
        await Promise.all([
          loadStudents(user.uid),
          loadSettings(user.uid)
        ]);
      } else {
        setAllStudents([]);
        setFilteredStudents([]);
        setSheets([]);
        setMappingSuccess(false);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const loadStudents = async (userId: string) => {
    try {
      const q = query(collection(db, "students"), where("userId", "==", userId));
      const snap = await getDocs(q);
      const loaded: any[] = [];
      const shSet = new Set<string>();
      
      snap.forEach(doc => {
        const data = doc.data();
        if(data.userId === userId) {
            loaded.push(data);
            if(data.sheet) shSet.add(data.sheet);
        }
      });
      if(loaded.length > 0) {
        setAllStudents(loaded);
        setFilteredStudents(loaded);
        setSheets(Array.from(shSet));
        setSelectedSheets(Array.from(shSet));
        setMappingSuccess(true);
        updateStats(loaded, Array.from(shSet).length);
      }
    } catch (err) {
      console.error(err);
    }
  };

  const loadSettings = async (userId: string) => {
    try {
      const docRef = doc(db, "settings", userId);
      const docSnap = await getDoc(docRef);
      if (docSnap.exists()) {
        const data = docSnap.data();
        if(data.aDir) setaDir(data.aDir);
        if(data.aSch) setaSch(data.aSch);
        if(data.aTeach) setaTeach(data.aTeach);
      }
    } catch(err) { console.error(err); }
  };

  const saveSettings = async () => {
    if(!user) return;
    setIsSavingSettings(true);
    try {
        await setDoc(doc(db, "settings", user.uid), {
            userId: user.uid,
            aDir,
            aSch,
            aTeach
        }, { merge: true });
        setShowSettingsModal(false);
    } catch (err: any) {
        alert('خطأ في حفظ الإعدادات: ' + err.message);
    }
    setIsSavingSettings(false);
  };

  const updateStats = (data: any[], shCount: number) => {
    const mal = data.filter(x => isMale(x.gender)).length;
    setStats({ tot: data.length, sh: shCount, mal, fem: data.length - mal });
  };

  const isMale = (g: any) => ['م', 'ذ', 'ذكر', 'm', 'male'].includes(String(g).toLowerCase().trim());

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  };

  const processFile = (file: File) => {
    if (!user) {
        alert("يرجى تسجيل الدخول أولاً لحفظ البيانات.");
        return;
    }
    setFileName(file.name);
    setIsSaving(true);
    setMappingSuccess(false);

    const rd = new FileReader();
    rd.onload = async (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target?.result as ArrayBuffer), { type: 'array' });
        let parsedAll: any[] = [];
        let parsedSheets = wb.SheetNames;
        wb.SheetNames.forEach(name => {
          XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', range: 10 }).forEach((r: any) => {
            if (!r || r.every((c: any) => c === '' || c == null)) return;
            const getVal = (idx: number) => idx >= 0 && idx < r.length ? String(r[idx] || '').trim() : '';
            parsedAll.push({
              code: getVal(colMap.code),
              lastname: getVal(colMap.lastname),
              firstname: getVal(colMap.firstname),
              gender: getVal(colMap.gender),
              dept: getVal(colMap.dept),
              sheet: name,
              userId: user.uid
            });
          });
        });

        if (parsedAll.length === 0) {
            alert('الملف فارغ أو لا توجد بيانات للتحميل بدءًا من الصف المحدد.');
            setIsSaving(false);
            return;
        }

        const batchSize = 400;
        let idCounter = 1;

        const q = query(collection(db, "students"), where("userId", "==", user.uid));
        const currSnap = await getDocs(q);
        
        let delBatch = writeBatch(db);
        let delCount = 0;
        for (const d of currSnap.docs) {
            if (d.data().userId === user.uid) {
                delBatch.delete(d.ref);
                delCount++;
                if (delCount === batchSize) {
                    await delBatch.commit();
                    delBatch = writeBatch(db);
                    delCount = 0;
                }
            }
        }
        if (delCount > 0) await delBatch.commit();

        let batch = writeBatch(db);
        let count = 0;

        for (let s of parsedAll) {
            s.id = idCounter++;
            const docRef = doc(collection(db, "students"));
            batch.set(docRef, s);
            count++;
            if(count === batchSize) {
                await batch.commit();
                batch = writeBatch(db);
                count = 0;
            }
        }
        if(count > 0) await batch.commit();

        setAllStudents(parsedAll);
        setFilteredStudents(parsedAll);
        setSheets(parsedSheets);
        setSelectedSheets(parsedSheets);
        updateStats(parsedAll, parsedSheets.length);
        setMappingSuccess(true);
        setActiveTab('students');
        setIsSaving(false);
      } catch (err: any) { alert('خطأ في قراءة وتحميل الملف:\n' + err.message); setIsSaving(false); }
    };
    rd.readAsArrayBuffer(file);
  };

  const deleteAllStudents = async () => {
    if (!user) return;
    
    setIsSaving(true);
    setShowDeleteConfirm(false);
    try {
        const q = query(collection(db, "students"), where("userId", "==", user.uid));
        const currSnap = await getDocs(q);
        
        let delBatch = writeBatch(db);
        let delCount = 0;
        for (const d of currSnap.docs) {
            delBatch.delete(d.ref);
            delCount++;
            if (delCount === 400) {
                await delBatch.commit();
                delBatch = writeBatch(db);
                delCount = 0;
            }
        }
        if (delCount > 0) await delBatch.commit();

        setAllStudents([]);
        setFilteredStudents([]);
        setSheets([]);
        setSelectedSheets([]);
        setStats({ tot: 0, sh: 0, mal: 0, fem: 0 });
        setMappingSuccess(false);
        setIsSaving(false);
        // Custom alert to not block UI if possible, or just a small notification, but standard alert is fine for now if it works, 
        // though native alert might also be blocked. Let's just rely on the UI changes
    } catch (err: any) {
        alert('حدث خطأ أثناء الحذف: ' + err.message);
        setIsSaving(false);
    }
  };

  useEffect(() => {
    const t = search.trim().toLowerCase();
    const sh = filterSheet;
    const gn = filterGender;
    const f = allStudents.filter(x => {
      const ms = !t || [x.lastname, x.firstname, x.code, x.dept].some(v => String(v).toLowerCase().includes(t));
      return ms && (!sh || x.sheet === sh) && (!gn || (gn === 'm' && isMale(x.gender)) || (gn === 'f' && !isMale(x.gender)));
    });
    setFilteredStudents(f);
  }, [search, filterSheet, filterGender, allStudents]);

  const toggleSheetSelection = (sh: string) => {
    setSelectedSheets(prev => prev.includes(sh) ? prev.filter(s => s !== sh) : [...prev, sh]);
  };
  const toggleAllSheets = (select: boolean) => {
    setSelectedSheets(select ? [...sheets] : []);
  };

  // --- PDF & Html Builders omitted for brevity, adding below ---
  return (
    <>
      <div className="particles" id="particles">
          {Array.from({length:15}).map((_, i) => (
             <div key={i} className="particle" style={{left: `${Math.random()*100}%`, animationDelay: `${Math.random()*20}s`, animationDuration: `${15+Math.random()*10}s`}}></div>
          ))}
      </div>

      <div className="hdr no-p">
        <div>
          <h1>نظام إدارة قوائم الطلاب</h1>
          <p>استيراد ملفات Excel - عرض البيانات الموحدة - ورقة الغياب النصف أسبوعية</p>
        </div>
        <div style={{display:'flex', alignItems:'center', gap:'16px', position: 'relative', zIndex: 10}}>
            <style>{`
                .hdr-btn-settings {
                    background: rgba(255,255,255,0.1);
                    border: none;
                    color: white;
                    cursor: pointer;
                    padding: 12px;
                    border-radius: 12px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    transition: all 0.2s ease;
                }
                .hdr-btn-settings:hover {
                    background: rgba(255,255,255,0.25);
                    transform: scale(1.05);
                }
                .hdr-btn-settings:active {
                    transform: scale(0.95);
                }
                .hdr-btn-logout {
                    background: none;
                    border: none;
                    color: white;
                    cursor: pointer;
                    padding: 6px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    opacity: 0.8;
                    transition: all 0.2s ease;
                }
                .hdr-btn-logout:hover {
                    opacity: 1;
                    transform: scale(1.05);
                }
                .hdr-btn-logout:active {
                    transform: scale(0.95);
                }
            `}</style>
            {user ? (
                <>
                    <button className="hdr-btn-settings" onClick={() => setShowSettingsModal(true)} title="الإعدادات">
                        <Settings size={22} />
                    </button>
                    <div style={{display:'flex', alignItems:'center', gap:'12px', background:'rgba(255,255,255,0.1)', padding:'8px 16px', borderRadius:'12px'}}>
                        <img src={user.photoURL || ''} alt="avatar" style={{width: 32, height: 32, borderRadius: '50%'}} />
                        <div style={{display:'flex', flexDirection:'column'}}>
                            <span style={{fontSize:'14px', fontWeight: 700}}>{user.displayName}</span>
                            <span style={{fontSize:'10px', opacity: 0.8}}>{user.email}</span>
                        </div>
                        <div style={{marginInlineStart: '12px', paddingInlineStart: '12px', borderInlineStart: '1px solid rgba(255,255,255,0.2)'}}>
                            <button className="hdr-btn-logout" onClick={logout} title="تسجيل الخروج">
                                <LogOut size={20} />
                            </button>
                        </div>
                    </div>
                </>
            ) : null}
            <span className="hdr-logo">🏫</span>
        </div>
      </div>

      <div className="wrap">
        {!user && !loading && (
             <div className="card" style={{padding: '40px', textAlign: 'center'}}>
                 <h2 style={{fontSize: '24px', color: 'var(--pri)', marginBottom: '16px', fontWeight: 800}}>أهلاً بك في نظام إدارة الغيابات</h2>
                 <p style={{color: '#64748b', marginBottom: '32px'}}>يرجى تسجيل الدخول باستخدام حساب Google للوصول إلى النظام وحفظ بيانات طلابك.</p>
                 <button onClick={loginWithGoogle} className="btn bs" style={{margin: '0 auto', fontSize: '16px', padding: '14px 32px'}}>
                     <LogIn size={20} /> تسجيل الدخول بواسطة Google
                 </button>
             </div>
        )}

        {loading && <div style={{textAlign:'center', padding:'40px', fontWeight:'bold', color:'var(--pri)'}}>جاري التحميل...</div>}

        {user && !loading && (
          <>
            <div className="card no-p">
              <div className="card-head">
                <h2>استيراد ملف Excel</h2>
                <small style={{color:'#64748b', fontSize:'12px'}}>يتم قراءة البيانات من جميع الأوراق تلقائياً</small>
              </div>
              <div className="card-body">
                {isSaving ? (
                    <div style={{textAlign:'center', padding:'40px', fontWeight:'bold', color:'var(--pri)'}}>جاري معالجة ورفع الملف السحابي... يرجى الانتظار</div>
                ) : !mappingSuccess && (
                    <>
                        <div className="up-zone" onDragOver={handleDrop} onDrop={handleDrop} onClick={() => document.getElementById('fileInput')?.click()}>
                        <span className="up-zone-ico">📂</span>
                        <h3>اضغط لاختيار الملف أو اسحبه هنا</h3>
                        <p>الصيغ المدعومة: .xlsx - .xls</p>
                        <p style={{color:'#94a3b8', fontSize:'13px', marginTop:'6px'}}>يُدمج محتوى جميع الأوراق في جدول واحد</p>
                        </div>
                        <input type="file" id="fileInput" accept=".xlsx,.xls" onChange={handleFileUpload} />
                    </>
                )}
                
                {mappingSuccess && (
                    <div className="ok-msg" style={{display:'flex'}}>
                        <span style={{fontSize:'20px'}}>✓</span> تم استيراد <strong>{sheets.length}</strong> قسم — <strong>{allStudents.length}</strong> طالب وتم حفظها سحابياً.
                    </div>
                )}
              </div>
            </div>

            {mappingSuccess && (
                <>
                    <div className="stats no-p">
                        <div className="stat"><div className="stat-n">{stats.tot}</div><div className="stat-l">إجمالي الطلاب</div></div>
                        <div className="stat"><div className="stat-n">{stats.sh}</div><div className="stat-l">عدد الأوراق</div></div>
                        <div className="stat"><div className="stat-n">{stats.mal}</div><div className="stat-l">ذكور</div></div>
                        <div className="stat"><div className="stat-n">{stats.fem}</div><div className="stat-l">إناث</div></div>
                    </div>

                    <div className="card">
                        <div className="tabs-nav no-p">
                            <button className={`t-btn ${activeTab === 'students' ? 'on' : ''}`} onClick={() => setActiveTab('students')}>قائمة الطلاب</button>
                            <button className={`t-btn ${activeTab === 'absence' ? 'on' : ''}`} onClick={() => setActiveTab('absence')}>ورقة الغياب النصف أسبوعية</button>
                        </div>

                        {activeTab === 'students' && (
                            <div className="t-pane on">
                                <div className="toolbar no-p">
                                    <input className="srch" placeholder="ابحث بالاسم أو الرمز أو القسم..." value={search} onChange={e => setSearch(e.target.value)} />
                                    <select className="filt" value={filterSheet} onChange={e => setFilterSheet(e.target.value)}>
                                        <option value="">كل الأقسام</option>
                                        {sheets.map(sh => <option key={sh} value={sh}>{sh}</option>)}
                                    </select>
                                    <select className="filt" value={filterGender} onChange={e => setFilterGender(e.target.value)}>
                                        <option value="">الجنس</option>
                                        <option value="m">ذكر</option>
                                        <option value="f">أنثى</option>
                                    </select>
                                    <button className="btn bs" onClick={() => {
                                        const ws = XLSX.utils.json_to_sheet(filteredStudents.map((x,i)=>({'#':i+1,'الرمز':x.code,'اللقب':x.lastname,'الاسم':x.firstname,'الجنس':x.gender,'تاريخ الازدياد':x.dept,'القسم':x.sheet})));
                                        const wb = XLSX.utils.book_new();
                                        XLSX.utils.book_append_sheet(wb, ws, "الطلاب");
                                        XLSX.writeFile(wb, "قائمة_الطلاب.xlsx");
                                    }}>تصدير Excel</button>
                                    <button className="btn bd" onClick={() => setShowDeleteConfirm(true)} disabled={isSaving}>
                                        {isSaving ? "جاري الحذف..." : "حذف الكل"}
                                    </button>
                                    <button className="btn bi" onClick={() => {
                                        const w = window.open('', '_blank');
                                        if(!w) return;
                                        w.document.write(`<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><title>قائمة الطلاب</title><style>body{font-family:Tajawal,Arial,sans-serif;direction:rtl;font-size:10px}table{width:100%;border-collapse:collapse}th{background:#0f172a;color:#fff;padding:6px;border:1px solid #333}td{padding:4px;border:1px solid #999;text-align:center}tr:nth-child(even){background:#f5f5f5}@page{margin:5mm;size:A4 portrait}</style></head><body>${document.querySelector('.dtbl')?.outerHTML}<script>window.onload=()=>window.print()</script></body></html>`);
                                        w.document.close();
                                    }}>طباعة القائمة</button>
                                </div>
                                <div className="tscroll">
                                    <table className="dtbl">
                                        <thead>
                                            <tr><th>#</th><th>الرمز</th><th>اللقب</th><th>الاسم</th><th>الجنس</th><th>تاريخ الازدياد</th><th>القسم</th></tr>
                                        </thead>
                                        <tbody>
                                            {filteredStudents.map((x, i) => (
                                                <tr key={x.id}>
                                                    <td>{i+1}</td>
                                                    <td><strong>{x.code}</strong></td>
                                                    <td>{x.lastname}</td>
                                                    <td>{x.firstname}</td>
                                                    <td className={isMale(x.gender)?'gm':'gf'}>{isMale(x.gender)?'ذكر':'أنثى'}</td>
                                                    <td>{x.dept}</td>
                                                    <td><span className="bsht">{x.sheet}</span></td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                                <p style={{color:'#64748b', fontSize:'12px', marginTop:'12px', textAlign:'center'}}>إجمالي: {filteredStudents.length}</p>
                            </div>
                        )}

                        {activeTab === 'absence' && <AbsenceView sheets={sheets} allStudents={allStudents} rawState={{aDir, setaDir, aSch, setaSch, aTri, setaTri, aFrom, setaFrom, aTo, setaTo, aTeach, setaTeach, aHalf, setaHalf, selectedSheets, setSelectedSheets, toggleSheetSelection, toggleAllSheets }} />}

                    </div>
                </>
            )}
          </>
        )}
      </div>

      {/* Settings Modal */}
      {showSettingsModal && (
        <div style={{position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:999, display:'flex', alignItems:'center', justifyContent:'center'}}>
            <div className="card" style={{width:'400px', padding:'24px', margin:0}}>
                <h3 style={{marginBottom:'16px', color:'var(--pri)', fontSize:'18px', fontWeight:700}}>إعدادات المؤسسة</h3>
                <div className="fg" style={{marginBottom:'12px'}}>
                    <label className="fl">المديرية</label>
                    <input className="fc" value={aDir} onChange={e=>setaDir(e.target.value)} placeholder="مثال: مديرية مراكش" />
                </div>
                <div className="fg" style={{marginBottom:'12px'}}>
                    <label className="fl">المؤسسة</label>
                    <input className="fc" value={aSch} onChange={e=>setaSch(e.target.value)} placeholder="اسم المؤسسة" />
                </div>
                <div className="fg" style={{marginBottom:'24px'}}>
                    <label className="fl">المربي / المربية</label>
                    <input className="fc" value={aTeach} onChange={e=>setaTeach(e.target.value)} placeholder="اسم الأستاذ(ة)" />
                </div>
                <div style={{display:'flex', gap:'10px', justifyContent:'flex-end'}}>
                    <button className="btn bl" onClick={()=>setShowSettingsModal(false)}>إلغاء</button>
                    <button className="btn bs" onClick={saveSettings} disabled={isSavingSettings}>{isSavingSettings ? 'جاري الحفظ...' : 'حفظ البيانات'}</button>
                </div>
            </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {showDeleteConfirm && (
        <div style={{position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:999, display:'flex', alignItems:'center', justifyContent:'center'}}>
            <div className="card" style={{width:'400px', padding:'24px', margin:0, textAlign: 'center'}}>
                <h3 style={{marginBottom:'16px', color:'#ef4444', fontSize:'18px', fontWeight:700}}>تأكيد الحذف</h3>
                <p style={{color:'#64748b', marginBottom:'24px'}}>هل أنت متأكد من أنك تريد حذف جميع بيانات الطلاب؟ لا يمكن التراجع عن هذه العملية.</p>
                <div style={{display:'flex', gap:'10px', justifyContent:'center'}}>
                    <button className="btn bl" onClick={()=>setShowDeleteConfirm(false)}>إلغاء</button>
                    <button className="btn bd" onClick={deleteAllStudents}>نعم، احذف الكل</button>
                </div>
            </div>
        </div>
      )}
    </>
  );
}

const HALF = [['الاثنين','الثلاثاء','الأربعاء'],['الخميس','الجمعة','السبت']];

function AbsenceView({sheets, allStudents, rawState}: any) {
    const {aDir, setaDir, aSch, setaSch, aTri, setaTri, aFrom, setaFrom, aTo, setaTo, aTeach, setaTeach, aHalf, setaHalf, selectedSheets, toggleSheetSelection, toggleAllSheets} = rawState;
    const [absHtml, setAbsHtml] = useState<string>('');
    const [built, setBuilt] = useState(false);

    // Provide the original behavior directly in the DOM using a ref and simple state updates so the classes act identical.
    const containerRef = React.useRef<HTMLDivElement>(null);

    const fdt = (d: string) => d ? new Date(d).toLocaleDateString('ar-DZ') : '........';

    const buildAbs = () => {
        if(selectedSheets.length === 0) { alert('حدد قسم واحد على الأقل'); return; }
        const days = HALF[aHalf];
        let fullHtml = '';

        selectedSheets.forEach((sheetName: string, gIdx: number) => {
            const students = allStudents.filter((s:any) => s.sheet === sheetName);
            if(!students.length) return;

            let th = `<thead><tr><th rowspan="3" class="amh">ر.ت</th><th rowspan="3" class="amh" style="width:110px">اللقب والاسم</th>`;
            days.forEach((d: string) => th += `<th colspan="8" class="adh day-separator">${d}</th>`);
            th += `</tr><tr>`;
            days.forEach(() => th += `<th colspan="4" class="asm day-separator">ص</th><th colspan="4" class="apm">م</th>`);
            th += `</tr><tr>`;
            days.forEach(() => {
                for(let h=1; h<=4; h++){
                    let cls = "ahm"; if(h===1) cls+=" day-separator";
                    th+=`<th class="${cls}">${h}</th>`;
                }
                for(let h=1; h<=4; h++){ th+=`<th class="ahp">${h}</th>`; }
            });
            th += `</tr></thead>`;

            let tb = `<tbody>`;
            students.forEach((x:any, i: number) => {
                const globalIdx = allStudents.indexOf(x);
                tb += `<tr><td class="ai">${i+1}</td><td class="ai name-cell" style="text-align:right;padding-right:4px;font-weight:600;white-space:normal;word-wrap:break-word;line-height:1.2">${x.lastname} ${x.firstname}</td>`;
                days.forEach((_:any, di:number) => {
                    for(let ses=0; ses<2; ses++){
                        for(let h=1; h<=4; h++){
                            let cls = "ac";
                            if(ses===0 && h===1) cls+=" day-separator";
                            tb+=`<td class="${cls}" data-ac="1" id="c_${globalIdx}_${di}_${ses}_${h}"></td>`;
                        }
                    }
                });
                tb+=`</tr>`;
            });
            tb += `</tbody>`;

            let tf = `<tfoot><tr class="sig-row"><td colspan="2" class="sig-label">توقيع الأستاذ (ة)</td>`;
            days.forEach(()=> {
                for(let ses=0; ses<2; ses++)
                    for(let h=1; h<=4; h++){
                        let cls = ""; if(ses===0 && h===1) cls = "day-separator";
                        tf += `<td class="${cls}"></td>`;
                    }
            });
            tf += `</tr><tr class="ref-row"><td class="ref-title"></td><td style="font-size:8pt;color:#666;text-align:center;vertical-align:middle">توجيه إلى الإدارة</td>`;
            days.forEach((d:string) => tf += `<td colspan="8" class="day-separator" style="text-align:center;padding:0"><div class="ref-day"></div></td>`);
            tf += `</tr></tfoot>`;

            const ph = `<div class="p-header p-only"><table style="width:100%;border:none;font-size:10pt;margin-bottom:2mm"><tr>
                <td style="border:none;text-align:right;vertical-align:top;width:35%">الأكاديميةالجهويةللتربيةوالتكوين<br>مراكش- أسفي<br>المديرية: <strong>${aDir||'....'}</strong><br>المؤسسة: <strong>${aSch||'....'}</strong></td>
                <td style="border:none;text-align:center;vertical-align:middle;width:30%"><div style="font-size:14pt;font-weight:bold;border:2px solid #333;padding:5px 15px;display:inline-block">ورقة الغياب النصف أسبوعية</div><br><strong style="font-size:12pt"></strong></td>
                <td style="border:none;text-align:left;vertical-align:top;width:35%">القسم: <strong style="font-size:12pt;color:#0369a1">${sheetName}</strong><br>من: ${fdt(aFrom)}<br>إلى: ${fdt(aTo)}<br> <strong></strong></td>
                </tr></table></div>`;
            const footer = `<div class="p-footer p-only" style="display:flex;justify-content:space-between;font-size:9pt"><span></span><span></span></div>`;
            const pageBreak = gIdx < selectedSheets.length-1 ? '<div class="page-break"></div>' : '';

            fullHtml += `<div class="abs-page" id="absPrint_${gIdx}">${ph}<div class="tbl-wrap"><table id="abt_${gIdx}">${th}${tb}${tf}</table></div>${footer}</div>${pageBreak}`;
        });

        setAbsHtml(fullHtml);
        setBuilt(true);
    };

    // Cycle Absence logic via Effect attaching to container
    useEffect(() => {
        if(!built || !containerRef.current) return;
        const STATES = ['','sg-abs','sg-jus','sg-lat'];
        const LABELS = ['','G','GE','T'];
        
        const handleClick = (e: MouseEvent) => {
            const t = e.target as HTMLElement;
            if(t.tagName === 'TD' && t.hasAttribute('data-ac')) {
                let cur = 0;
                STATES.forEach((st, i) => { if(st && t.classList.contains(st)) cur = i; });
                STATES.forEach(st => { if(st) t.classList.remove(st); });
                const nxt = (cur + 1) % STATES.length;
                if(STATES[nxt]) t.classList.add(STATES[nxt]);
                t.textContent = LABELS[nxt];
            }
        };

        const cnt = containerRef.current;
        cnt.addEventListener('click', handleClick);
        return () => cnt.removeEventListener('click', handleClick);
    }, [built]);


    const clearAbs = () => {
        if(!confirm('مسح جميع الغيابات؟')) return;
        if(containerRef.current) {
            containerRef.current.querySelectorAll('.ac').forEach(c => {
                c.classList.remove('sg-abs', 'sg-jus', 'sg-lat');
                c.textContent = '';
            });
        }
    };

    const downloadPdf = async () => {
        const pages = containerRef.current?.querySelectorAll('.abs-page');
        if(!pages || !pages.length) { alert('⚠️ أنشئ ورقة الغياب أولاً'); return; }
        const btn = document.getElementById('btnPdf') as HTMLButtonElement;
        btn.innerHTML = 'جاري التحميل...'; btn.disabled=true;
        try{
            const pdf = new jsPDF({orientation:'p', unit:'mm', format:'a4'});
            document.body.classList.add('generating-pdf');
            for(let i=0; i<pages.length; i++) {
                const page = pages[i] as HTMLElement; 
                page.style.display = 'flex';
                // Remove React strict mode double-renders causing issues by using raw DOM wait
                await new Promise(r => setTimeout(r, 100)); 
                const canvas = await html2canvas(page, {scale: 2, useCORS: true, logging: false, width: page.offsetWidth, height: page.offsetHeight});
                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                const pdfWidth = pdf.internal.pageSize.getWidth();
                const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
                let finalHeight = pdfHeight > pdf.internal.pageSize.getHeight() ? pdf.internal.pageSize.getHeight() : pdfHeight;
                if(i > 0) pdf.addPage();
                pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, finalHeight);
            }
            document.body.classList.remove('generating-pdf');
            pdf.save('ورقة_الغياب.pdf');
        }catch(err: any){
            alert('حدث خطأ'); console.error(err); document.body.classList.remove('generating-pdf');
        }
        btn.innerHTML = 'تحميل PDF'; btn.disabled=false;
    };

    const printAbs = () => {
        const pages = containerRef.current?.querySelectorAll('.abs-page');
        if(!pages || !pages.length) { alert('⚠️ أنشئ ورقة الغياب أولاً'); return; }
        let allHtml = '';
        pages.forEach((p, i) => { allHtml += p.outerHTML; if(i < pages.length-1) allHtml += '<div style="page-break-after:always"></div>'; });
        const w = window.open('', '_blank');
        if(!w) return;
        w.document.write(`<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><title>ورقة الغياب</title><style>
        html,body{height:100%;margin:0;padding:0}
        body{font-family:Tajawal,Arial,sans-serif;direction:rtl}
        .abs-page{display:flex;flex-direction:column;height:100%;width:100%}
        .p-header{flex-shrink:0}
        .tbl-wrap{flex:1 1 auto;display:flex;flex-direction:column;width:100%;min-height:0}
        .tbl-wrap table{flex:1;width:100%;height:100%;border-collapse:collapse;table-layout:fixed;font-size:8pt}
        thead{height:auto}
        thead th{padding:1px 2px !important; vertical-align:middle}
        tbody{height:100%}
        tbody tr{height:1px}
        .tbl-wrap th,.tbl-wrap td{border:1px solid #333;padding:0;margin:0;vertical-align:middle;text-align:center}
        .name-cell{white-space:normal;word-wrap:break-word;text-align:right;line-height:1.1}
        .p-footer{flex-shrink:0;margin-top:3mm}
        .amh{background:#0f172a!important;color:#fff!important;font-size:9pt!important;font-weight:bold!important}
        .adh{background:#fff!important;color:#000!important;font-size:8pt!important;font-weight:900!important}
        .asm{background:#bfdbfe!important;color:#1e3a8a!important;font-weight:bold!important;font-size:7pt!important}
        .apm{background:#fed7aa!important;color:#9a3412!important;font-weight:bold!important;font-size:7pt!important}
        .ahm{background:#dcfce7!important;font-size:6pt!important;font-weight:bold!important}
        .ahp{background:#fef3c7!important;font-size:6pt!important;font-weight:bold!important}
        .ai{background:#f8fafc!important;font-weight:600!important;font-size:7pt!important}
        .sg-abs{background:#ef4444!important;color:#fff!important;font-weight:bold}
        .sg-jus{background:#f59e0b!important;color:#fff!important;font-weight:bold}
        .sg-lat{background:#a855f7!important;color:#fff!important;font-weight:bold}
        tfoot .sig-row td{height:40px!important}
        tfoot .ref-row td{height:50px!important}
        .day-separator { border-right: 4px solid #333 !important; }
        @page{margin:8mm;size:A4 portrait}
        </style></head><body>${allHtml}<script>window.onload=()=>window.print()</script></body></html>`);
        w.document.close();
    };

    const expAbsXls = () => {
        const tables = containerRef.current?.querySelectorAll('[id^="abt_"]');
        if(!tables || !tables.length) { alert('⚠️ أنشئ ورقة الغياب أولاً'); return; }
        const wb = XLSX.utils.book_new();
        tables.forEach((t, i) => {
            const ws = XLSX.utils.table_to_sheet(t);
            XLSX.utils.book_append_sheet(wb, ws, (selectedSheets[i] || `ورقة${i+1}`).substring(0,31));
        });
        XLSX.writeFile(wb, 'ورقة_الغياب.xlsx');
    };

    return (
        <div className="t-pane on">
            <div className="abs-set no-p">
                <h3>إعدادات ورقة الغياب النصف أسبوعية</h3>
                <div className="sg">
                    <div className="fg"><label className="fl">المديرية</label><input className="fc" value={aDir} onChange={e=>setaDir(e.target.value)} placeholder="مديرية التربية" /></div>
                    <div className="fg"><label className="fl">المؤسسة</label><input className="fc" value={aSch} onChange={e=>setaSch(e.target.value)} placeholder="اسم المؤسسة" /></div>
                    <div className="fg"><label className="fl">الثلاثي</label>
                        <select className="fc" value={aTri} onChange={e=>setaTri(e.target.value)}><option>الثلاثي الأول</option><option>الثلاثي الثاني</option><option>الثلاثي الثالث</option></select>
                    </div>
                    <div className="fg"><label className="fl">من تاريخ</label><input className="fc" type="date" value={aFrom} onChange={e=>setaFrom(e.target.value)} /></div>
                    <div className="fg"><label className="fl">إلى تاريخ</label><input className="fc" type="date" value={aTo} onChange={e=>setaTo(e.target.value)} /></div>
                    <div className="fg"><label className="fl">المربي / المربية</label><input className="fc" value={aTeach} onChange={e=>setaTeach(e.target.value)} placeholder="اسم المربي/ة" /></div>
                    <div className="fg"><label className="fl">نصف الأسبوع</label>
                        <select className="fc" value={aHalf} onChange={e=>setaHalf(parseInt(e.target.value))}>
                            <option value="0">الاثنين - الثلاثاء - الأربعاء</option>
                            <option value="1">الخميس - الجمعة - السبت</option>
                        </select>
                    </div>

                    <div className="fg sg-wide">
                        <label className="fl">اختيار الأقسام للطباعة (كل قسم في صفحة مستقلة)</label>
                        <div className="selector-box">
                            {sheets.length === 0 ? <div className="selector-empty">استورد ملف Excel أولاً</div> : (
                                <>
                                    <div className="selector-toolbar">
                                        <button className="bs" onClick={()=>toggleAllSheets(true)} style={{fontSize:'11px'}}>تحديد الكل</button>
                                        <button className="bl" onClick={()=>toggleAllSheets(false)} style={{fontSize:'11px'}}>إلغاء الكل</button>
                                        <span style={{color:'#64748b', fontSize:'11px', marginRight:'auto'}}>{sheets.length} أقسام</span>
                                    </div>
                                    {sheets.map((sh:string) => {
                                        const count = allStudents.filter((s:any) => s.sheet === sh).length;
                                        return (
                                            <div className="selector-item" key={sh}>
                                                <input type="checkbox" id={`sh_${sh}`} checked={selectedSheets.includes(sh)} onChange={()=>toggleSheetSelection(sh)} />
                                                <label htmlFor={`sh_${sh}`}>{sh}</label>
                                                <span className="count">{count}</span>
                                            </div>
                                        )
                                    })}
                                </>
                            )}
                        </div>
                    </div>
                </div>

                <div style={{display:'flex', gap:'10px', flexWrap:'wrap'}}>
                    <button className="btn bp" onClick={buildAbs}>إنشاء ورقة الغياب</button>
                    <button className="btn bd" onClick={printAbs}>طباعة</button>
                    <button className="btn bi" id="btnPdf" onClick={downloadPdf}>تحميل PDF</button>
                    <button className="btn ba" onClick={clearAbs}>مسح الغيابات</button>
                    <button className="btn bl" onClick={expAbsXls}>تصدير Excel</button>
                </div>
            </div>

            {built && (
                <div className="lgnd no-p">
                    <strong>المفتاح (انقر على الخلية للتبديل):</strong>
                    <div className="li"><div className="ld" style={{background:'linear-gradient(135deg,#ef4444,#f87171)'}}></div>غائب - G</div>
                    <div className="li"><div className="ld" style={{background:'linear-gradient(135deg,#f59e0b,#fbbf24)'}}></div>غائب بعذر - GE</div>
                    <div className="li"><div className="ld" style={{background:'linear-gradient(135deg,#a855f7,#c084fc)'}}></div>متأخر - T</div>
                    <div className="li"><div className="ld" style={{background:'#fff', border:'1px solid #cbd5e1'}}></div>حاضر (مسح)</div>
                </div>
            )}

            <div ref={containerRef}>
                {!built ? (
                    <div className="empty"><span className="empty-ico">📝</span><p>استورد ملف Excel ثم اضغط "إنشاء ورقة الغياب"</p></div>
                ) : (
                    <div dangerouslySetInnerHTML={{__html: absHtml}}></div>
                )}
            </div>
        </div>
    )
}
