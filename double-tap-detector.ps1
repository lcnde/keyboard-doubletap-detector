<#
    Double Tap Detector
    -------------------
    Controlla se la tastiera "batte due volte" lo stesso tasto (double tapping / chattering).

    Come funziona
      - Legge la tastiera in modo passivo con la Raw Input API di Windows (sola lettura): non blocca,
        non modifica e non ritarda nessun tasto, non serve essere amministratore, non tocca il registro.
      - Soglie tarate per utente che scrive a 120wpm (record: 70 ms tra due pressioni con un dito, 43 ms con
        due dita; nessuna ripressione sotto i 20 ms dal rilascio):
          DOPPIO TAP   lo stesso tasto premuto di nuovo entro 35 ms dalla pressione precedente (sotto il
                       record umano con due dita) oppure entro 10 ms dal suo rilascio (rimbalzo al rilascio).
                       Notifica + log.
          SOSPETTO     di nuovo premuto tra 35 e 60 ms dopo la pressione precedente: sotto il minimo con un
                       dito, quindi impossibile scrivendo normalmente. Solo log e statistiche, niente notifica.
      - I tasti premuti normalmente NON vengono salvati da nessuna parte: nel log finiscono solo doppi tap
        e pressioni sospette.
      - Gli input simulati dai programmi (password manager, macro, tastiera su schermo) vengono ignorati.
      - Gli intervalli si misurano con il timer ad alta precisione; se il PC era rallentato e la misura
        non è affidabile, l'evento viene scartato invece di dare un falso allarme.

    File
      log\doppi-tap.csv   una riga per ogni doppio tap o pressione sospetta (si apre con Excel)
      log\sessioni.log    avvio e arresto del monitoraggio, con riepilogo

    Finestra:  Q = esci   H = riduci a icona   L = apri la cartella dei log   R = azzera le statistiche
    Icona nell'area di notifica: clic = mostra la finestra, tasto destro = menu (log, esci)

    Parametri
      -SogliaMs 35            doppio tap: ms massimi tra due pressioni dello stesso tasto
      -SogliaRilascioMs 10    doppio tap: ms massimi tra il rilascio e la nuova pressione
      -SogliaSospettoMs 60    sospetto: ms massimi tra due pressioni
      -PausaNotificheSec 30   tempo minimo tra due notifiche (i doppi tap nel frattempo vengono sommati)
      -Autotest               verifica che tutto funzioni e termina (non scrive nei log)
#>
param(
    [double]$SogliaMs = 35,
    [double]$SogliaRilascioMs = 10,
    [double]$SogliaSospettoMs = 60,
    [int]$PausaNotificheSec = 30,
    [switch]$Autotest
)

$ErrorActionPreference = 'Stop'

$cartellaLog  = Join-Path $PSScriptRoot 'log'
$fileCsv      = Join-Path $cartellaLog 'doppi-tap.csv'
$fileSessioni = Join-Path $cartellaLog 'sessioni.log'
$utf8         = New-Object System.Text.UTF8Encoding($true)   # con BOM, così Excel legge bene gli accenti
$sepCsv       = (Get-Culture).TextInfo.ListSeparator         # lo stesso separatore che usa Excel su questo PC
$testoSoglia  = $SogliaMs.ToString('0.##')
$testoSogliaR = $SogliaRilascioMs.ToString('0.##')
$testoSogliaS = $SogliaSospettoMs.ToString('0.##')

$codiceCSharp = @'
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RilevatoreDoppioTap
{
    public class Rilevamento
    {
        public DateTime Ora;
        public string Tasto;
        public string Tastiera;
        public string Tipo;
        public double DaPressioneMs;   // dalla pressione precedente dello stesso tasto
        public double DaRilascioMs;    // dal rilascio precedente (-1 = nessun rilascio in mezzo)
        public double IntervalloMs;    // l'intervallo che ha fatto scattare il rilevamento
        public string Riferimento;     // "dopo la pressione precedente" / "dopo il rilascio"
        public bool Sospetto;          // true = pressione sospetta (niente notifica), false = doppio tap
    }

    public class StatTasto
    {
        public string Tasto;
        public long Pressioni;
        public long DoppiTap;
        public long Sospetti;
        public double MinMs = -1;          // intervallo minimo tra due pressioni
        public double MinRilascioMs = -1;  // intervallo minimo tra un rilascio e la pressione successiva
        public double UltimoMs = -1;       // intervallo tra le ultime due pressioni
        public StatTasto Copia() { return (StatTasto)MemberwiseClone(); }
    }

    public class Istantanea
    {
        public DateTime Avvio;
        public long Pressioni, DoppiTap, Sospetti, Scartati, SimulatiIgnorati, Errori;
        public string UltimoTasto;
        public double UltimoMs, UltimoRilascioMs;
        public StatTasto[] Tasti;
    }

    class StatoTasto
    {
        public bool Premuto, HaPressione, HaRilascio;
        public long TPressione, TRilascio, TUltimoEvento;
        public int MsgPressione, MsgRilascio;
    }

    internal delegate bool GestoreCtrl(int tipo);

    static class Nativo
    {
        public const uint RID_INPUT = 0x10000003;
        public const uint RIDI_DEVICENAME = 0x20000007;
        public const int RIM_TYPEKEYBOARD = 1;
        public const uint RIDEV_INPUTSINK = 0x00000100;
        public const int WM_SETICON = 0x0080;

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWINPUTDEVICE { public ushort usUsagePage; public ushort usUsage; public uint dwFlags; public IntPtr hwndTarget; }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Explicit)]
        public struct INPUTUNION { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT { public uint type; public INPUTUNION u; }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterRawInputDevices(RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);
        [DllImport("user32.dll")]
        public static extern uint GetRawInputData(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetRawInputDeviceInfoW")]
        public static extern uint GetRawInputDeviceInfo(IntPtr hDevice, uint uiCommand, StringBuilder pData, ref uint pcbSize);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetKeyNameTextW")]
        public static extern int GetKeyNameText(int lParam, StringBuilder lpString, int cchSize);
        [DllImport("user32.dll")] public static extern int GetMessageTime();
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
        [DllImport("kernel32.dll")] public static extern bool SetConsoleCtrlHandler(GestoreCtrl handler, bool add);
    }

    // Finestra invisibile che riceve i messaggi WM_INPUT della tastiera
    class FinestraInput : NativeWindow
    {
        public FinestraInput() { CreateHandle(new CreateParams()); }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Rilevatore.WM_INPUT)
            {
                long ts = Stopwatch.GetTimestamp();   // subito, per la massima precisione
                Rilevatore.LeggiInput(m.LParam, ts, Nativo.GetMessageTime());
            }
            else if (m.Msg == Rilevatore.WM_AZIONE)
            {
                Rilevatore.EseguiAzioni();
                return;
            }
            else if (m.Msg == Rilevatore.WM_ENDSESSION && m.WParam != IntPtr.Zero)
            {
                Rilevatore.Chiusura("arresto o disconnessione di Windows");
            }
            base.WndProc(ref m);
        }
    }

    public static class Rilevatore
    {
        internal const int WM_INPUT = 0x00FF;
        internal const int WM_ENDSESSION = 0x0016;
        internal const int WM_AZIONE = 0x8001;   // WM_APP + 1

        public static double SogliaMs = 35;           // doppio tap: pressione -> pressione
        public static double SogliaRilascioMs = 10;   // doppio tap: rilascio -> pressione (rimbalzo al rilascio)
        public static double SogliaSospettoMs = 60;   // sospetto: pressione -> pressione
        public static bool IncludiInputSimulato = false;
        public static volatile bool UscitaRichiesta = false;
        public static string FileSessioni;
        public static string CartellaLog;

        static readonly object _lock = new object();
        static readonly Dictionary<string, StatoTasto> _stati = new Dictionary<string, StatoTasto>();
        static readonly Dictionary<string, StatTasto> _stat = new Dictionary<string, StatTasto>();
        static readonly ConcurrentQueue<Rilevamento> _rilevamenti = new ConcurrentQueue<Rilevamento>();
        static readonly ConcurrentQueue<Action> _azioni = new ConcurrentQueue<Action>();
        static readonly Dictionary<string, string> _nomiTasti = new Dictionary<string, string>();
        static readonly Dictionary<IntPtr, string> _tastiere = new Dictionary<IntPtr, string>();
        static long _pressioni, _doppi, _sospetti, _scartati, _simulati, _errori;
        static string _ultimoTasto = "";
        static double _ultimoMs = -1, _ultimoRilascioMs = -1;
        static DateTime _avvio = DateTime.Now;
        static int _riepilogoScritto;

        static Thread _thread;
        static FinestraInput _finestra;
        static NotifyIcon _icona;
        static string _erroreAvvio;
        static IntPtr _buffer = IntPtr.Zero;
        static uint _dimBuffer;
        static GestoreCtrl _gestoreCtrl;   // tenuto in un campo statico perché il GC non lo elimini

        static double Ms(long ticks) { return ticks * 1000.0 / Stopwatch.Frequency; }

        // ---------------------------------------------------------------- logica di rilevamento

        public static void Elabora(string chiave, string tasto, string tastiera, bool rilascio, long ts, int oraMsg)
        {
            Rilevamento ril = null;
            lock (_lock)
            {
                StatoTasto s;
                if (!_stati.TryGetValue(chiave, out s)) { s = new StatoTasto(); _stati[chiave] = s; }

                if (rilascio)
                {
                    if (s.Premuto) { s.Premuto = false; s.HaRilascio = true; s.TRilascio = ts; s.MsgRilascio = oraMsg; }
                    s.TUltimoEvento = ts;
                    return;
                }

                double daPressione = s.HaPressione ? Ms(ts - s.TPressione) : -1;
                double daRilascio = -1;
                string tipo = null;
                bool sospetto = false, dalRilascio = false;
                double limite = 0;
                int msgRiferimento = 0;

                if (s.Premuto)
                {
                    // Nessun rilascio ricevuto dall'ultima pressione
                    double daUltimo = Ms(ts - s.TUltimoEvento);
                    s.TUltimoEvento = ts;
                    if (daPressione >= 0 && daPressione <= SogliaSospettoMs)
                    {
                        // Troppo presto per la ripetizione automatica (parte dopo almeno 250 ms): segnale doppio
                        sospetto = daPressione > SogliaMs;
                        tipo = sospetto ? "sospetto (senza rilascio)" : "doppia pressione senza rilascio";
                        limite = sospetto ? SogliaSospettoMs : SogliaMs;
                        msgRiferimento = s.MsgPressione;
                    }
                    else if (daUltimo < 1500)
                    {
                        return;   // ripetizione automatica del tasto tenuto giu': non e' un nuovo tap
                    }
                    // altrimenti il rilascio e' andato perso (es. sospensione del PC): conta come nuova pressione
                }
                else
                {
                    if (s.HaRilascio && s.HaPressione && s.TRilascio >= s.TPressione) daRilascio = Ms(ts - s.TRilascio);
                    if (daPressione >= 0 && daPressione <= SogliaMs)
                    {
                        tipo = "doppia pressione"; limite = SogliaMs; msgRiferimento = s.MsgPressione;
                    }
                    else if (daRilascio >= 0 && daRilascio <= SogliaRilascioMs)
                    {
                        tipo = "ripressione subito dopo il rilascio"; dalRilascio = true; limite = SogliaRilascioMs; msgRiferimento = s.MsgRilascio;
                    }
                    else if (daPressione >= 0 && daPressione <= SogliaSospettoMs)
                    {
                        tipo = "sospetto"; sospetto = true; limite = SogliaSospettoMs; msgRiferimento = s.MsgPressione;
                    }
                }

                bool scartato = false;
                if (tipo != null && unchecked(oraMsg - msgRiferimento) > limite + 20)
                {
                    // Per l'orologio di Windows i due eventi erano lontani: il PC era rallentato e la misura
                    // precisa non e' affidabile. Meglio scartare che dare un falso allarme.
                    scartato = true;
                    tipo = null;
                    _scartati++;
                }

                s.Premuto = true; s.HaPressione = true; s.TPressione = ts; s.MsgPressione = oraMsg; s.TUltimoEvento = ts;

                StatTasto st;
                if (!_stat.TryGetValue(tasto, out st)) { st = new StatTasto(); st.Tasto = tasto; _stat[tasto] = st; }
                st.Pressioni++;
                _pressioni++;
                st.UltimoMs = scartato ? -1 : daPressione;
                if (!scartato && daPressione >= 0 && (st.MinMs < 0 || daPressione < st.MinMs)) st.MinMs = daPressione;
                if (!scartato && daRilascio >= 0 && (st.MinRilascioMs < 0 || daRilascio < st.MinRilascioMs)) st.MinRilascioMs = daRilascio;
                _ultimoTasto = tasto;
                _ultimoMs = st.UltimoMs;
                _ultimoRilascioMs = scartato ? -1 : daRilascio;

                if (tipo != null)
                {
                    if (sospetto) { st.Sospetti++; _sospetti++; }
                    else { st.DoppiTap++; _doppi++; }
                    ril = new Rilevamento();
                    ril.Sospetto = sospetto;
                    ril.Ora = DateTime.Now;
                    ril.Tasto = tasto;
                    ril.Tastiera = tastiera;
                    ril.Tipo = tipo;
                    ril.DaPressioneMs = daPressione;
                    ril.DaRilascioMs = daRilascio;
                    ril.IntervalloMs = dalRilascio ? daRilascio : daPressione;
                    ril.Riferimento = dalRilascio ? "dopo il rilascio" : "dopo la pressione precedente";
                }
            }
            if (ril != null) _rilevamenti.Enqueue(ril);
        }

        public static bool ProssimoRilevamento(out Rilevamento r) { return _rilevamenti.TryDequeue(out r); }

        public static Istantanea Stato()
        {
            var s = new Istantanea();
            var tasti = new List<StatTasto>();
            lock (_lock)
            {
                s.Avvio = _avvio; s.Pressioni = _pressioni; s.DoppiTap = _doppi; s.Sospetti = _sospetti; s.Scartati = _scartati;
                s.SimulatiIgnorati = _simulati; s.UltimoTasto = _ultimoTasto; s.UltimoMs = _ultimoMs; s.UltimoRilascioMs = _ultimoRilascioMs;
                foreach (StatTasto t in _stat.Values) tasti.Add(t.Copia());
            }
            s.Errori = Interlocked.Read(ref _errori);
            tasti.Sort(delegate(StatTasto a, StatTasto b)
            {
                int c = b.DoppiTap.CompareTo(a.DoppiTap);
                if (c != 0) return c;
                c = b.Sospetti.CompareTo(a.Sospetti);
                if (c != 0) return c;
                c = b.Pressioni.CompareTo(a.Pressioni);
                if (c != 0) return c;
                return string.Compare(a.Tasto, b.Tasto, StringComparison.CurrentCulture);
            });
            s.Tasti = tasti.ToArray();
            return s;
        }

        public static void Azzera()
        {
            lock (_lock)
            {
                _stat.Clear();
                _pressioni = 0; _doppi = 0; _sospetti = 0; _scartati = 0; _simulati = 0;
                _ultimoTasto = ""; _ultimoMs = -1; _ultimoRilascioMs = -1;
            }
        }

        // ---------------------------------------------------------------- lettura della tastiera (Raw Input)

        internal static void LeggiInput(IntPtr hRaw, long ts, int oraMsg)
        {
            try
            {
                uint dimHeader = (uint)(8 + 2 * IntPtr.Size);
                uint dim = 0;
                if (Nativo.GetRawInputData(hRaw, Nativo.RID_INPUT, IntPtr.Zero, ref dim, dimHeader) != 0 || dim == 0) return;
                if (dim > _dimBuffer)
                {
                    if (_buffer != IntPtr.Zero) Marshal.FreeHGlobal(_buffer);
                    _buffer = Marshal.AllocHGlobal((int)dim);
                    _dimBuffer = dim;
                }
                uint letti = Nativo.GetRawInputData(hRaw, Nativo.RID_INPUT, _buffer, ref dim, dimHeader);
                if (letti == uint.MaxValue || letti < dimHeader + 8) return;
                if (Marshal.ReadInt32(_buffer, 0) != Nativo.RIM_TYPEKEYBOARD) return;

                IntPtr dispositivo = Marshal.ReadIntPtr(_buffer, 8);
                int o = (int)dimHeader;
                ushort makeCode = (ushort)Marshal.ReadInt16(_buffer, o);
                ushort flags = (ushort)Marshal.ReadInt16(_buffer, o + 2);
                ushort vkey = (ushort)Marshal.ReadInt16(_buffer, o + 6);

                // Input generato da un programma (SendInput): nessun dispositivo fisico
                if (dispositivo == IntPtr.Zero && !IncludiInputSimulato)
                {
                    lock (_lock) _simulati++;
                    return;
                }
                // Codici "finti" che Windows aggiunge ad alcuni tasti (es. Maiusc virtuale, Pausa)
                if (vkey == 0xFF || makeCode == 0xFF) return;

                bool rilascio = (flags & 1) != 0;
                bool e0 = (flags & 2) != 0;
                bool e1 = (flags & 4) != 0;
                string codice = (e1 ? "E1-" : e0 ? "E0-" : "") + makeCode.ToString("X2") + (makeCode == 0 ? "-" + vkey.ToString("X2") : "");
                string chiave = dispositivo.ToInt64().ToString("X") + "/" + codice;
                Elabora(chiave, NomeTasto(codice, makeCode, e0, e1, vkey), NomeTastiera(dispositivo), rilascio, ts, oraMsg);
            }
            catch
            {
                Interlocked.Increment(ref _errori);
            }
        }

        static string NomeTasto(string codice, ushort makeCode, bool e0, bool e1, ushort vkey)
        {
            string nome;
            if (_nomiTasti.TryGetValue(codice, out nome)) return nome;
            nome = null;
            if (!e1 && makeCode != 0)
            {
                var sb = new StringBuilder(64);
                if (Nativo.GetKeyNameText((makeCode << 16) | (e0 ? (1 << 24) : 0), sb, 64) > 0) nome = sb.ToString();
            }
            if (string.IsNullOrEmpty(nome)) nome = ((Keys)vkey).ToString();
            _nomiTasti[codice] = nome;
            return nome;
        }

        static string NomeTastiera(IntPtr dispositivo)
        {
            if (dispositivo == IntPtr.Zero) return "input simulato";
            string nome;
            if (_tastiere.TryGetValue(dispositivo, out nome)) return nome;
            nome = "";
            try
            {
                uint dim = 0;
                Nativo.GetRawInputDeviceInfo(dispositivo, Nativo.RIDI_DEVICENAME, null, ref dim);
                if (dim > 0)
                {
                    var sb = new StringBuilder((int)dim + 1);
                    if ((int)Nativo.GetRawInputDeviceInfo(dispositivo, Nativo.RIDI_DEVICENAME, sb, ref dim) > 0) nome = sb.ToString();
                }
            }
            catch { }
            // "\\?\HID#VID_046D&PID_C31C&MI_00#7&2a...#{guid}"  ->  "HID#VID_046D&PID_C31C&MI_00"
            if (nome.StartsWith(@"\\?\")) nome = nome.Substring(4);
            string[] parti = nome.Split('#');
            if (parti.Length >= 2) nome = parti[0] + "#" + parti[1];
            if (nome.Length == 0) nome = "tastiera " + dispositivo.ToInt64().ToString("X");
            _tastiere[dispositivo] = nome;
            return nome;
        }

        // ---------------------------------------------------------------- avvio / arresto

        public static void Avvia(bool conIcona)
        {
            if (_thread != null) return;
            _avvio = DateTime.Now;
            var pronto = new ManualResetEvent(false);
            _thread = new Thread(delegate()
            {
                try
                {
                    Application.ThreadException += delegate(object s, ThreadExceptionEventArgs e) { Interlocked.Increment(ref _errori); };
                    _finestra = new FinestraInput();
                    var dispositivi = new Nativo.RAWINPUTDEVICE[1];
                    dispositivi[0].usUsagePage = 0x01;                  // Generic Desktop
                    dispositivi[0].usUsage = 0x06;                      // tastiera
                    dispositivi[0].dwFlags = Nativo.RIDEV_INPUTSINK;    // ricevi anche quando non e' in primo piano
                    dispositivi[0].hwndTarget = _finestra.Handle;
                    if (!Nativo.RegisterRawInputDevices(dispositivi, 1, (uint)Marshal.SizeOf(typeof(Nativo.RAWINPUTDEVICE))))
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    if (conIcona) CreaIcona();
                }
                catch (Exception ex)
                {
                    _erroreAvvio = ex.Message;
                    pronto.Set();
                    return;
                }
                pronto.Set();
                Application.Run();
                if (_icona != null) { _icona.Visible = false; _icona.Dispose(); _icona = null; }
                _finestra.DestroyHandle();
                _finestra = null;
            });
            _thread.IsBackground = true;
            _thread.Priority = ThreadPriority.Highest;   // passa quasi tutto il tempo in attesa: serve solo a misurare con precisione
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            pronto.WaitOne(15000);
            if (_erroreAvvio != null) throw new InvalidOperationException("Impossibile leggere la tastiera: " + _erroreAvvio);
        }

        public static void Ferma()
        {
            if (_thread == null) return;
            InAccoda(delegate() { Application.ExitThread(); });
            _thread.Join(3000);
            _thread = null;
        }

        public static void RegistraChiusuraFinestra()
        {
            _gestoreCtrl = new GestoreCtrl(delegate(int tipo)
            {
                if (tipo == 2) Chiusura("finestra chiusa");   // CTRL_CLOSE_EVENT
                else if (tipo != 0) ScriviSessione("Segnale di chiusura dalla console (tipo " + tipo + ")");
                return false;
            });
            Nativo.SetConsoleCtrlHandler(_gestoreCtrl, true);
            AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
            {
                ScriviSessione("Errore imprevisto: " + e.ExceptionObject);
                Chiusura("errore imprevisto");
            };
        }

        internal static void Chiusura(string motivo)
        {
            ScriviRiepilogo(motivo);
            // l'icona va tolta dal suo thread: si aspetta al massimo 1 secondo
            var fatto = new ManualResetEvent(false);
            InAccoda(delegate() { if (_icona != null) _icona.Visible = false; fatto.Set(); });
            fatto.WaitOne(1000);
        }

        public static void ScriviSessione(string testo)
        {
            if (string.IsNullOrEmpty(FileSessioni)) return;
            try { File.AppendAllText(FileSessioni, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + testo + Environment.NewLine, new UTF8Encoding(true)); }
            catch { }
        }

        public static void ScriviRiepilogo(string motivo)
        {
            if (Interlocked.Exchange(ref _riepilogoScritto, 1) == 1) return;
            if (string.IsNullOrEmpty(FileSessioni)) return;
            try
            {
                Istantanea s = Stato();
                TimeSpan d = DateTime.Now - s.Avvio;
                var tasti = new List<string>();
                foreach (StatTasto t in s.Tasti) if (t.DoppiTap > 0) tasti.Add(t.Tasto + " x" + t.DoppiTap);
                string riga = string.Format("{0:yyyy-MM-dd HH:mm:ss}  Arresto ({1}) - durata {2}:{3:00}:{4:00}, pressioni {5}, doppi tap {6}{7}, sospetti {8}",
                    DateTime.Now, motivo, (int)d.TotalHours, d.Minutes, d.Seconds, s.Pressioni, s.DoppiTap,
                    tasti.Count > 0 ? " (" + string.Join(", ", tasti.ToArray()) + ")" : "", s.Sospetti);
                File.AppendAllText(FileSessioni, riga + Environment.NewLine, new UTF8Encoding(true));
            }
            catch { }
        }

        // ---------------------------------------------------------------- icona, notifiche, finestra

        static void InAccoda(Action azione)
        {
            _azioni.Enqueue(azione);
            FinestraInput f = _finestra;
            if (f != null && f.Handle != IntPtr.Zero) Nativo.PostMessage(f.Handle, WM_AZIONE, IntPtr.Zero, IntPtr.Zero);
        }

        internal static void EseguiAzioni()
        {
            Action a;
            while (_azioni.TryDequeue(out a))
            {
                try { a(); } catch { Interlocked.Increment(ref _errori); }
            }
        }

        static void CreaIcona()
        {
            Icon immagine = DisegnaIcona();
            _icona = new NotifyIcon();
            _icona.Icon = immagine;
            _icona.Text = "Double Tap Detector";
            var menu = new ContextMenuStrip();
            menu.Items.Add("Mostra i dati", null, delegate(object s, EventArgs e) { MostraConsole(); });
            menu.Items.Add("Apri la cartella dei log", null, delegate(object s, EventArgs e)
            {
                if (!string.IsNullOrEmpty(CartellaLog)) Process.Start("explorer.exe", "\"" + CartellaLog + "\"");
            });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Esci", null, delegate(object s, EventArgs e) { UscitaRichiesta = true; });
            _icona.ContextMenuStrip = menu;
            _icona.MouseClick += delegate(object s, MouseEventArgs e) { if (e.Button == MouseButtons.Left) MostraConsole(); };
            _icona.Visible = true;

            // stessa icona anche per la finestra nella barra delle applicazioni
            IntPtr console = Nativo.GetConsoleWindow();
            if (console != IntPtr.Zero)
            {
                Nativo.PostMessage(console, Nativo.WM_SETICON, IntPtr.Zero, immagine.Handle);
                Nativo.PostMessage(console, Nativo.WM_SETICON, new IntPtr(1), immagine.Handle);
            }
        }

        static Icon DisegnaIcona()
        {
            using (var bmp = new Bitmap(32, 32))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                using (var percorso = new GraphicsPath())
                using (var sfondo = new SolidBrush(Color.FromArgb(214, 64, 52)))
                using (var carattere = new Font("Segoe UI", 15f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var formato = new StringFormat())
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    g.Clear(Color.Transparent);
                    const int d = 12;
                    percorso.AddArc(1, 1, d, d, 180, 90);
                    percorso.AddArc(30 - d, 1, d, d, 270, 90);
                    percorso.AddArc(30 - d, 30 - d, d, d, 0, 90);
                    percorso.AddArc(1, 30 - d, d, d, 90, 90);
                    percorso.CloseFigure();
                    g.FillPath(sfondo, percorso);
                    formato.Alignment = StringAlignment.Center;
                    formato.LineAlignment = StringAlignment.Center;
                    g.DrawString("2x", carattere, Brushes.White, new RectangleF(0, 1, 32, 32), formato);
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }

        public static void Notifica(string titolo, string testo, bool avviso)
        {
            InAccoda(delegate()
            {
                if (_icona != null) _icona.ShowBalloonTip(10000, titolo, testo, avviso ? ToolTipIcon.Warning : ToolTipIcon.Info);
            });
        }

        public static void ImpostaTooltip(string testo)
        {
            if (testo.Length > 63) testo = testo.Substring(0, 63);
            InAccoda(delegate() { if (_icona != null) _icona.Text = testo; });
        }

        public static void MostraFinestra(IntPtr h)
        {
            if (h == IntPtr.Zero) return;
            Nativo.ShowWindow(h, Nativo.IsIconic(h) ? 9 : 5);   // SW_RESTORE / SW_SHOW
            Nativo.SetForegroundWindow(h);
        }

        public static void MostraConsole() { MostraFinestra(Nativo.GetConsoleWindow()); }

        public static void RiduciConsole()
        {
            IntPtr h = Nativo.GetConsoleWindow();
            if (h != IntPtr.Zero) Nativo.ShowWindow(h, 6);   // SW_MINIMIZE
        }

        public static bool ConsoleRidotta()
        {
            IntPtr h = Nativo.GetConsoleWindow();
            return h != IntPtr.Zero && Nativo.IsIconic(h);
        }

        // ---------------------------------------------------------------- solo per l'autotest

        public static int InviaTastoTest(ushort vk, ushort scan, int volte)
        {
            var input = new Nativo.INPUT[volte * 2];
            for (int i = 0; i < input.Length; i++)
            {
                input[i].type = 1;   // INPUT_KEYBOARD
                input[i].u.ki.wVk = vk;
                input[i].u.ki.wScan = scan;
                input[i].u.ki.dwFlags = (uint)(i % 2 == 1 ? 2 : 0);   // KEYEVENTF_KEYUP per i rilasci
            }
            return (int)Nativo.SendInput((uint)input.Length, input, Marshal.SizeOf(typeof(Nativo.INPUT)));
        }
    }
}
'@

# ==================================================================== funzioni

function Testo-Ms([double]$ms) {
    if ($ms -lt 0) { return '-' }
    return '{0:0.0} ms' -f $ms
}

function Riga([string]$testo, [string]$colore = 'Gray') {
    [pscustomobject]@{ Testo = $testo; Colore = $colore }
}

function Scrivi-Sessione([string]$testo) {
    try { [IO.File]::AppendAllText($fileSessioni, ('{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $testo) + "`r`n", $utf8) } catch { }
}

function Riga-Csv($r) {
    $campi = @(
        $r.Ora.ToString('yyyy-MM-dd'), $r.Ora.ToString('HH:mm:ss.fff'), $r.Tasto,
        $(if ($r.DaPressioneMs -ge 0) { $r.DaPressioneMs.ToString('0.0') } else { '' }),
        $(if ($r.DaRilascioMs -ge 0) { $r.DaRilascioMs.ToString('0.0') } else { '' }),
        $r.Tipo, $r.Tastiera)
    ($campi | ForEach-Object { '"' + ([string]$_).Replace('"', '""') + '"' }) -join $sepCsv
}

function Svuota-Log {
    if ($daScrivere.Count -eq 0) { return }
    $testo = (($daScrivere | ForEach-Object { Riga-Csv $_ }) -join "`r`n") + "`r`n"
    try {
        [IO.File]::AppendAllText($fileCsv, $testo, $utf8)
        $daScrivere.Clear()
        $script:logBloccato = $false
    } catch {
        $script:logBloccato = $true   # file aperto in un altro programma (es. Excel): si riprova al giro dopo
    }
}

function Prepara-Console {
    try { [Console]::CursorVisible = $false } catch { }
    try {
        $l = [Math]::Min(124, [Console]::LargestWindowWidth)
        $a = [Math]::Min(40, [Console]::LargestWindowHeight)
        [Console]::SetBufferSize([Math]::Max($l, [Console]::BufferWidth), [Math]::Max($a, [Console]::BufferHeight))
        [Console]::SetWindowPosition(0, 0)
        [Console]::SetWindowSize($l, $a)
        [Console]::SetBufferSize($l, $a)
    } catch { }
}

function Disegna {
    $s = $DTD::Stato()
    $larg = [Console]::WindowWidth
    $alt  = [Console]::WindowHeight
    if ("$larg x $alt" -ne $script:dimensioni) {
        try { [Console]::SetWindowPosition(0, 0); [Console]::SetBufferSize($larg, $alt) } catch { }
        [Console]::Clear()
        $script:dimensioni = "$larg x $alt"
    }
    $sep = ' ' + ('-' * [Math]::Max(10, $larg - 3))
    $righe = New-Object System.Collections.Generic.List[object]

    $durata = (Get-Date) - $s.Avvio
    $righe.Add((Riga (' DOUBLE TAP DETECTOR    in ascolto da {0}:{1:00}:{2:00}    (avviato il {3:dd/MM} alle {3:HH:mm})' -f [int][Math]::Floor($durata.TotalHours), $durata.Minutes, $durata.Seconds, $s.Avvio) 'Cyan'))
    $righe.Add((Riga $sep 'DarkGray'))
    $righe.Add((Riga " Soglie:  DOPPIO TAP = ripremuto entro $testoSoglia ms dalla pressione o $testoSogliaR ms dal rilascio    SOSPETTO = entro $testoSogliaS ms"))
    $perc = ''
    if ($s.Pressioni -gt 0) { $perc = ' ({0:0.00}%)' -f (100.0 * $s.DoppiTap / $s.Pressioni) }
    $righe.Add((Riga (' Pressioni: {0:N0}     Doppi tap: {1:N0}{2}     Sospetti: {3:N0}' -f $s.Pressioni, $s.DoppiTap, $perc, $s.Sospetti)))
    $righe.Add((Riga (' Ignorati: {0:N0} input simulati da programmi, {1:N0} misure non affidabili (PC rallentato)' -f $s.SimulatiIgnorati, $s.Scartati) 'DarkGray'))
    $righe.Add((Riga ''))
    if ($s.DoppiTap -gt 0) {
        $coinvolti = ($s.Tasti | Where-Object { $_.DoppiTap -gt 0 } | ForEach-Object { $_.Tasto + ' (' + $_.DoppiTap + ')' }) -join ', '
        $esito = (' ESITO: ATTENZIONE - {0:N0} doppi tap rilevati. Tasti: ' -f $s.DoppiTap) + $coinvolti
        $righe.Add((Riga $esito 'Red'))
    } elseif ($s.Sospetti -gt 0) {
        $coinvolti = ($s.Tasti | Where-Object { $_.Sospetti -gt 0 } | ForEach-Object { $_.Tasto + ' (' + $_.Sospetti + ')' }) -join ', '
        $righe.Add((Riga (" ESITO: nessun doppio tap. Sospetti ($testoSoglia-$testoSogliaS ms, normali solo se fatti apposta a due dita): " + $coinvolti) 'Yellow'))
    } else {
        $righe.Add((Riga ' ESITO: nessun doppio tap rilevato finora.' 'Green'))
    }
    if ($s.UltimoTasto) {
        $ultimo = ' Ultimo tasto premuto: ' + $s.UltimoTasto
        if ($s.UltimoMs -ge 0) {
            $ultimo += '   (dalla pressione precedente dello stesso tasto: ' + (Testo-Ms $s.UltimoMs)
            if ($s.UltimoRilascioMs -ge 0) { $ultimo += ', dal suo rilascio: ' + (Testo-Ms $s.UltimoRilascioMs) }
            $ultimo += ')'
        }
        $righe.Add((Riga $ultimo))
    } else {
        $righe.Add((Riga ' Ultimo tasto premuto: -'))
    }
    if ($script:logBloccato) { $righe.Add((Riga ' Il file di log è aperto in un altro programma: i dati verranno scritti appena lo chiudi.' 'Yellow')) }
    $righe.Add((Riga ''))

    $righe.Add((Riga ' ULTIMI DOPPI TAP (rosso) E SOSPETTI (giallo)' 'Yellow'))
    $righe.Add((Riga ('   {0,-12}  {1,-18} {2,16} {3,13}   {4,-36} {5}' -f 'Ora', 'Tasto', 'Tra le pressioni', 'Dal rilascio', 'Tipo', 'Tastiera') 'DarkGray'))
    if ($recenti.Count -eq 0) {
        $righe.Add((Riga '   (nessuno)' 'DarkGray'))
    } else {
        $n = [Math]::Min($recenti.Count, 6)
        for ($i = 0; $i -lt $n; $i++) {
            $r = $recenti[$i]
            $colore = 'Red'
            if ($r.Sospetto) { $colore = 'Yellow' }
            $righe.Add((Riga ('   {0,-12}  {1,-18} {2,16} {3,13}   {4,-36} {5}' -f $r.Ora.ToString('HH:mm:ss.fff'), $r.Tasto, (Testo-Ms $r.DaPressioneMs), (Testo-Ms $r.DaRilascioMs), $r.Tipo, $r.Tastiera) $colore))
        }
    }
    $righe.Add((Riga ''))

    $righe.Add((Riga ' STATISTICHE PER TASTO   (rosso = doppi tap, giallo = sospetti)' 'Yellow'))
    $righe.Add((Riga ('   {0,-18} {1,9} {2,9} {3,9} {4,17} {5,17} {6,18}' -f 'Tasto', 'Pressioni', 'Doppi tap', 'Sospetti', 'Min tra pressioni', 'Min dal rilascio', 'Ultimo intervallo') 'DarkGray'))
    $piede = 3
    $spazio = $alt - 1 - $righe.Count - $piede
    $tasti = $s.Tasti
    if ($tasti.Length -eq 0) {
        $righe.Add((Riga '   (premi qualche tasto)' 'DarkGray'))
    } else {
        $mostra = $tasti.Length
        if ($mostra -gt $spazio) { $mostra = [Math]::Max(0, $spazio - 1) }
        for ($i = 0; $i -lt $mostra; $i++) {
            $t = $tasti[$i]
            $colore = 'Gray'
            if ($t.DoppiTap -gt 0) { $colore = 'Red' } elseif ($t.Sospetti -gt 0) { $colore = 'Yellow' }
            $righe.Add((Riga ('   {0,-18} {1,9:N0} {2,9:N0} {3,9:N0} {4,17} {5,17} {6,18}' -f $t.Tasto, $t.Pressioni, $t.DoppiTap, $t.Sospetti, (Testo-Ms $t.MinMs), (Testo-Ms $t.MinRilascioMs), (Testo-Ms $t.UltimoMs)) $colore))
        }
        if ($mostra -lt $tasti.Length) { $righe.Add((Riga ('   ... e altri {0} tasti (ingrandisci la finestra per vederli tutti)' -f ($tasti.Length - $mostra)) 'DarkGray')) }
    }

    while ($righe.Count -lt $alt - 1 - $piede) { $righe.Add((Riga '')) }
    $righe.Add((Riga $sep 'DarkGray'))
    $righe.Add((Riga " Log: $fileCsv" 'DarkGray'))
    $righe.Add((Riga ' [Q] esci   [H] riduci a icona   [L] apri log   [R] azzera statistiche   (ridotta a icona continua a controllare)' 'DarkCyan'))

    for ($i = 0; $i -lt $alt - 1; $i++) {
        $testo = ''
        $colore = 'Gray'
        if ($i -lt $righe.Count) { $testo = $righe[$i].Testo; $colore = $righe[$i].Colore }
        if ($testo.Length -ge $larg) { $testo = $testo.Substring(0, $larg - 1) } else { $testo = $testo.PadRight($larg - 1) }
        [Console]::SetCursorPosition(0, $i)
        [Console]::ForegroundColor = [ConsoleColor]$colore
        [Console]::Write($testo)
    }
    [Console]::ResetColor()
}

function Invoke-Autotest {
    $T = [RilevatoreDoppioTap.Rilevatore]
    [RilevatoreDoppioTap.Rilevatore]::SogliaMs = 35
    [RilevatoreDoppioTap.Rilevatore]::SogliaRilascioMs = 10
    [RilevatoreDoppioTap.Rilevatore]::SogliaSospettoMs = 60
    $freq = [Diagnostics.Stopwatch]::Frequency / 1000.0
    $base = [Diagnostics.Stopwatch]::GetTimestamp()
    $script:falliti = 0

    # evento sintetico: tasto, 'giu'/'su', istante in ms (e, se serve, l'ora del messaggio di Windows in ms)
    function Ev([string]$tasto, [string]$azione, [double]$ms, [int]$msg = -1) {
        if ($msg -lt 0) { $msg = [int]$ms }
        $T::Elabora("test/$tasto", $tasto, 'test', ($azione -eq 'su'), $base + [long]($ms * $freq), $msg)
    }
    function Verifica([string]$descrizione, [string]$tasto, [long]$doppi, [long]$sospetti, [long]$pressioni) {
        $st = $T::Stato().Tasti | Where-Object { $_.Tasto -eq $tasto } | Select-Object -First 1
        if ($null -ne $st -and $st.DoppiTap -eq $doppi -and $st.Sospetti -eq $sospetti -and $st.Pressioni -eq $pressioni) {
            Write-Host "  OK      $descrizione" -ForegroundColor Green
        } else {
            Write-Host ("  ERRORE  {0}: attesi {1} doppi / {2} sospetti / {3} pressioni, trovati {4} / {5} / {6}" -f $descrizione, $doppi, $sospetti, $pressioni, $st.DoppiTap, $st.Sospetti, $st.Pressioni) -ForegroundColor Red
            $script:falliti++
        }
    }

    Write-Host 'Test della logica (eventi simulati; doppio tap 35 ms / 10 ms dal rilascio, sospetto 60 ms):'
    Ev A giu 0; Ev A su 80; Ev A giu 200; Ev A su 260
    Verifica 'Digitazione normale (stesso tasto dopo 200 ms)' A 0 0 2
    Ev U1 giu 0; Ev U1 su 35; Ev U1 giu 70; Ev U1 su 105
    Verifica 'Il tuo record con un dito (70 ms): nessun allarme' U1 0 0 2
    Ev U2 giu 0; Ev U2 su 25; Ev U2 giu 43; Ev U2 su 70
    Verifica 'Il tuo record con due dita (43 ms): solo sospetto' U2 0 1 2
    Ev B giu 0; Ev B su 5; Ev B giu 12; Ev B su 60
    Verifica 'Rimbalzo in pressione (12 ms)' B 1 0 2
    Ev B2 giu 0; Ev B2 su 15; Ev B2 giu 30; Ev B2 su 80
    Verifica 'Rimbalzo a 30 ms (sfuggiva alla vecchia soglia di 20 ms)' B2 1 0 2
    Ev C giu 0; Ev C su 90; Ev C giu 95; Ev C su 97
    Verifica 'Rimbalzo al rilascio (5 ms dopo il rilascio)' C 1 0 2
    Ev D giu 0; Ev D giu 500; Ev D giu 533; Ev D giu 566; Ev D su 600
    Verifica 'Tasto tenuto premuto (ripetizione automatica)' D 0 0 1
    Ev E giu 0; Ev E giu 8; Ev E su 70
    Verifica 'Doppia pressione senza rilascio (8 ms)' E 1 0 2
    Ev F giu 0; Ev F su 20; Ev F giu 35; Ev F su 80
    Verifica 'Esattamente 35 ms tra le pressioni: doppio tap' F 1 0 2
    Ev G giu 0; Ev G su 20; Ev G giu 36; Ev G su 80
    Verifica '36 ms tra le pressioni: sospetto' G 0 1 2
    Ev S giu 0; Ev S su 30; Ev S giu 60; Ev S su 100
    Verifica 'Esattamente 60 ms tra le pressioni: sospetto' S 0 1 2
    Ev M giu 0; Ev M su 50; Ev M giu 60; Ev M su 100
    Verifica 'Esattamente 10 ms dal rilascio: doppio tap' M 1 0 2
    Ev N giu 0; Ev N su 50; Ev N giu 61; Ev N su 100
    Verifica '11 ms dal rilascio e 61 ms tra le pressioni: niente' N 0 0 2
    $scartatiPrima = $T::Stato().Scartati
    Ev H giu 0 0; Ev H su 50 50; Ev H giu 55 120
    Verifica 'PC rallentato: misura scartata, nessun falso allarme' H 0 0 2
    if ($T::Stato().Scartati -ne $scartatiPrima + 1) { Write-Host '  ERRORE  il contatore degli scartati non è aumentato' -ForegroundColor Red; $script:falliti++ }
    Ev I giu 0; Ev I giu 5000
    Verifica 'Rilascio perso, nuova pressione dopo 5 s' I 0 0 2
    Ev J giu 0; Ev K giu 5; Ev J su 50; Ev K su 60
    Verifica 'Tasti diversi ravvicinati (J)' J 0 0 1
    Verifica 'Tasti diversi ravvicinati (K)' K 0 0 1

    $doppi = 0; $sospetti = 0
    $r = $null
    while ($T::ProssimoRilevamento([ref]$r)) { if ($r.Sospetto) { $sospetti++ } else { $doppi++ } }
    if ($doppi -eq 6 -and $sospetti -eq 3) { Write-Host '  OK      in coda 6 doppi tap (notifica + log) e 3 sospetti (solo log)' -ForegroundColor Green }
    else { Write-Host "  ERRORE  in coda $doppi doppi tap e $sospetti sospetti invece di 6 e 3" -ForegroundColor Red; $script:falliti++ }

    Write-Host ''
    Write-Host 'Test con la lettura reale della tastiera di Windows (Raw Input):'
    [RilevatoreDoppioTap.Rilevatore]::IncludiInputSimulato = $true
    $prima = $T::Stato()
    try {
        $T::Avvia($false)
        Start-Sleep -Milliseconds 300
        $inviati = $T::InviaTastoTest(0x87, 0x76, 2)   # F24 due volte di fila: tasto assente sulle tastiere normali
        Start-Sleep -Milliseconds 700
        $dopo = $T::Stato()
    } finally {
        $T::Ferma()
    }
    $nuovi = $dopo.DoppiTap - $prima.DoppiTap
    if ($inviati -eq 4 -and $nuovi -ge 1) {
        Write-Host "  OK      F24 premuto 2 volte via Windows: letto e rilevato come doppio tap" -ForegroundColor Green
    } else {
        Write-Host "  ERRORE  eventi inviati: $inviati, doppi tap rilevati: $nuovi" -ForegroundColor Red
        $script:falliti++
    }

    Write-Host ''
    if ($script:falliti -eq 0) { Write-Host 'AUTOTEST SUPERATO' -ForegroundColor Green } else { Write-Host "AUTOTEST FALLITO ($($script:falliti) errori)" -ForegroundColor Red }
    return $script:falliti
}

# ==================================================================== avvio

try {
    if (-not ('RilevatoreDoppioTap.Rilevatore' -as [type])) {
        Add-Type -TypeDefinition $codiceCSharp -ReferencedAssemblies System.Windows.Forms, System.Drawing
    }
} catch {
    Write-Host "Errore nella preparazione del rilevatore: $($_.Exception.Message)" -ForegroundColor Red
    New-Item -ItemType Directory -Force -Path $cartellaLog | Out-Null
    Scrivi-Sessione "Errore all'avvio: $($_.Exception.Message)"
    try { [void][Console]::ReadKey($true) } catch { }
    exit 1
}
$DTD = [RilevatoreDoppioTap.Rilevatore]
[RilevatoreDoppioTap.Rilevatore]::SogliaMs = $SogliaMs
[RilevatoreDoppioTap.Rilevatore]::SogliaRilascioMs = $SogliaRilascioMs
[RilevatoreDoppioTap.Rilevatore]::SogliaSospettoMs = $SogliaSospettoMs

if ($Autotest) { exit (Invoke-Autotest) }

# Una sola istanza alla volta: se è già attivo, mostra la finestra esistente ed esci
$mutex = New-Object System.Threading.Mutex($false, 'Local\DoubleTapDetector')
$haMutex = $false
try { $haMutex = $mutex.WaitOne(0) }
catch {
    if ($_.Exception -is [System.Threading.AbandonedMutexException] -or $_.Exception.InnerException -is [System.Threading.AbandonedMutexException]) { $haMutex = $true }
    else { throw }
}
if (-not $haMutex) {
    Write-Host 'Double Tap Detector è già in esecuzione: lo trovi nella barra delle applicazioni.' -ForegroundColor Yellow
    Get-Process powershell -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -ne $PID -and $_.MainWindowTitle -like 'Double Tap Detector*' } |
        ForEach-Object { $DTD::MostraFinestra($_.MainWindowHandle) }
    Start-Sleep -Seconds 4
    exit 0
}

try {
    New-Item -ItemType Directory -Force -Path $cartellaLog | Out-Null
    if (-not (Test-Path $fileCsv)) {
        $intestazione = @('Data', 'Ora', 'Tasto', 'Tra le pressioni (ms)', 'Dal rilascio (ms)', 'Tipo', 'Tastiera') -join $sepCsv
        [IO.File]::WriteAllText($fileCsv, $intestazione + "`r`n", $utf8)
    }
    [RilevatoreDoppioTap.Rilevatore]::FileSessioni = $fileSessioni
    [RilevatoreDoppioTap.Rilevatore]::CartellaLog = $cartellaLog

    $Host.UI.RawUI.WindowTitle = 'Double Tap Detector'
    Prepara-Console
    $DTD::Avvia($true)
    $DTD::RegistraChiusuraFinestra()
    Scrivi-Sessione "Avvio (doppio tap: $testoSoglia ms tra le pressioni o $testoSogliaR ms dal rilascio; sospetto: $testoSogliaS ms)"
    $DTD::Notifica('Double Tap Detector attivo', "Ti avviso se un tasto viene registrato due volte (entro $testoSoglia ms, o $testoSogliaR ms dal rilascio). Clicca sull'icona per vedere i dati in tempo reale.", $false)
} catch {
    Scrivi-Sessione "Errore all'avvio: $($_.Exception.Message)"
    [Console]::ResetColor()
    Write-Host "Errore all'avvio: $($_.Exception.Message)" -ForegroundColor Red
    $DTD::MostraConsole()
    try { [void][Console]::ReadKey($true) } catch { }
    exit 1
}

$recenti        = New-Object System.Collections.Generic.List[object]
$daScrivere     = New-Object System.Collections.Generic.List[object]
$daNotificare   = 0
$ultimoDaNotif  = $null
$ultimaNotifica = [datetime]::MinValue
$ultimoTotale   = -1
$logBloccato    = $false
$dimensioni     = ''
$finestraPreparata = $false
$esci           = $false
$motivoUscita   = 'interrotto con Ctrl+C'

try {
    while (-not $esci) {
        if ($DTD::UscitaRichiesta) { $motivoUscita = "chiuso dal menu dell'icona"; break }

        # nuovi doppi tap e sospetti: in memoria e nel log; si notificano solo i doppi tap
        $r = $null
        while ($DTD::ProssimoRilevamento([ref]$r)) {
            $recenti.Insert(0, $r)
            if ($recenti.Count -gt 50) { $recenti.RemoveAt(50) }
            $daScrivere.Add($r)
            if (-not $r.Sospetto) {
                $daNotificare++
                $ultimoDaNotif = $r
            }
        }
        Svuota-Log

        if ($daNotificare -gt 0 -and ((Get-Date) - $ultimaNotifica).TotalSeconds -ge $PausaNotificheSec) {
            $u = $ultimoDaNotif
            $dettaglio = 'tasto ' + $u.Tasto + ', secondo tap ' + (Testo-Ms $u.IntervalloMs) + ' ' + $u.Riferimento
            if ($daNotificare -eq 1) { $testo = 'Doppio tap: ' + $dettaglio + '.' }
            else { $testo = "$daNotificare doppi tap. Ultimo: " + $dettaglio + '.' }
            $testo += ' Totale della sessione: ' + $DTD::Stato().DoppiTap + '.'
            $DTD::Notifica('Double tap rilevato', $testo, $true)
            $daNotificare = 0
            $ultimaNotifica = Get-Date
        }

        # titolo della finestra (visibile passando sulla barra delle applicazioni) e tooltip dell'icona
        $totale = $DTD::Stato().DoppiTap
        if ($totale -ne $ultimoTotale) {
            $ultimoTotale = $totale
            if ($totale -eq 0) { $stato = 'nessun doppio tap' } elseif ($totale -eq 1) { $stato = '1 doppio tap' } else { $stato = "$totale doppi tap" }
            $Host.UI.RawUI.WindowTitle = "Double Tap Detector - $stato"
            $DTD::ImpostaTooltip("Double Tap Detector - $stato")
        }

        try {
            while ([Console]::KeyAvailable) {
                $k = [Console]::ReadKey($true)
                switch ($k.Key) {
                    'Q' { $esci = $true; $motivoUscita = 'chiuso con Q' }
                    'H' { $DTD::RiduciConsole() }
                    'L' { Start-Process explorer.exe -ArgumentList "`"$cartellaLog`"" }
                    'R' { $DTD::Azzera(); $recenti.Clear(); $ultimoTotale = -1 }
                }
            }
        } catch { }

        # disegna solo se la finestra non è ridotta a icona (nessuno spreco quando è in background)
        if (-not $DTD::ConsoleRidotta()) {
            # partendo ridotta a icona Windows ignora le dimensioni: si impostano alla prima apertura
            if (-not $finestraPreparata) { Prepara-Console; $finestraPreparata = $true }
            try { Disegna } catch { $dimensioni = '' }
        }
        Start-Sleep -Milliseconds 200
    }
} catch {
    $motivoUscita = 'errore: ' + $_.Exception.Message
} finally {
    Svuota-Log
    $DTD::ScriviRiepilogo($motivoUscita)
    $DTD::Ferma()
    try { [Console]::ResetColor(); [Console]::CursorVisible = $true; [Console]::Clear() } catch { }
    try { $mutex.ReleaseMutex() } catch { }
}
