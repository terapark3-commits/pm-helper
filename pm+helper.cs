using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MssqlPatientHelper
{
    public class PasswordForm : Form
    {
        private TextBox _txtPassword;
        private Button _btnOk;
        private Button _btnCancel;
        private Label _lblMsg;

        public PasswordForm()
        {
            InitializeComponent();
            try
            {
                this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "보안 인증 | 만든이: 한솔인텍";
            this.Size = new Size(350, 180);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(15, 23, 42); // Slate 900
            this.ForeColor = Color.FromArgb(248, 250, 252); // Slate 50
            this.Font = new Font("맑은 고딕", 9.75F, FontStyle.Regular);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            _lblMsg = new Label
            {
                Text = "시스템 시작을 위해 비밀번호를 입력하십시오.",
                Location = new Point(20, 20),
                Size = new Size(300, 25),
                ForeColor = Color.FromArgb(148, 163, 184) // Slate 400
            };
            this.Controls.Add(_lblMsg);

            _txtPassword = new TextBox
            {
                Location = new Point(20, 50),
                Size = new Size(290, 25),
                PasswordChar = '*',
                BackColor = Color.FromArgb(30, 41, 59), // Slate 800
                ForeColor = Color.FromArgb(248, 250, 252),
                BorderStyle = BorderStyle.FixedSingle
            };
            _txtPassword.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    VerifyPassword();
                }
            };
            this.Controls.Add(_txtPassword);

            // Maker Label
            Label lblMaker = new Label
            {
                Text = "만든이: 한솔인텍",
                Location = new Point(20, 102),
                Size = new Size(100, 20),
                ForeColor = Color.FromArgb(100, 116, 139), // Slate 500 (dimmer text)
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            this.Controls.Add(lblMaker);

            _btnOk = new Button
            {
                Text = "확인",
                Location = new Point(130, 95),
                Size = new Size(85, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(99, 102, 241), // Indigo 500
                ForeColor = Color.White
            };
            _btnOk.FlatAppearance.BorderSize = 0;
            _btnOk.Click += (s, e) => VerifyPassword();
            this.Controls.Add(_btnOk);

            _btnCancel = new Button
            {
                Text = "취소",
                Location = new Point(225, 95),
                Size = new Size(85, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(51, 65, 85), // Slate 700
                ForeColor = Color.FromArgb(248, 250, 252)
            };
            _btnCancel.FlatAppearance.BorderSize = 0;
            _btnCancel.Click += (s, e) => {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            this.Controls.Add(_btnCancel);
        }

        private void VerifyPassword()
        {
            if (_txtPassword.Text == "a134679a**")
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("비밀번호가 올바르지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _txtPassword.Text = "";
                _txtPassword.Focus();
            }
        }
    }

    public class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => {
                Exception ex = e.ExceptionObject as Exception;
                string msg = ex != null ? ex.ToString() : "알 수 없는 시스템 예외";
                MessageBox.Show("치명적 시스템 오류가 발생했습니다:\n\n" + msg, "시스템 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            Application.ThreadException += (s, e) => {
                MessageBox.Show("프로그램 실행 중 오류가 발생했습니다:\n\n" + e.Exception.ToString(), "실행 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            try
            {
                using (PasswordForm pf = new PasswordForm())
                {
                    if (pf.ShowDialog() == DialogResult.OK)
                    {
                        Application.Run(new MainForm());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("MainForm 시작 중 예외 발생:\n\n" + ex.ToString(), "초기화 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    public class TablessTabControl : TabControl
    {
        private const int TCM_ADJUSTRECT = 0x1328;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == TCM_ADJUSTRECT)
            {
                m.Result = (IntPtr)1;
                return;
            }

            base.WndProc(ref m);
        }
    }

    public class MainForm : Form
    {
        private sealed class JuminRestoreTablePlan
        {
            public string TableName;
            public string DisplayName;
            public string ChartColumn;
            public string NameColumn;
            public string JuminColumn;
            public string RowKeyExpression;
            public int RowCount;
        }

        private sealed class JuminEncryptionRestorePlan
        {
            public string ChartNo;
            public string PatientName;
            public string JuminPrefix;
            public string OldCipher;
            public string BackupCipher;
            public string BackupDatabase;
            public List<JuminRestoreTablePlan> Tables = new List<JuminRestoreTablePlan>();

            public int TotalRows
            {
                get { return Tables.Sum(t => t.RowCount); }
            }
        }

        private sealed class JuminEncryptionRestoreResult
        {
            public Guid SessionId;
            public int UpdatedRows;
            public bool Committed;
        }

        // Dark Theme Colors
        private static readonly Color ColorBgMain = Color.FromArgb(15, 23, 42);       // Slate 900
        private static readonly Color ColorBgCard = Color.FromArgb(24, 33, 47);       // Fluent dark surface
        private static readonly Color ColorBorder = Color.FromArgb(58, 69, 88);       // Fluent border
        private static readonly Color ColorTextMain = Color.FromArgb(248, 250, 252);  // Slate 50
        private static readonly Color ColorTextSec = Color.FromArgb(173, 181, 194);   // Fluent secondary text
        private static readonly Color ColorIndigo = Color.FromArgb(0, 120, 212);      // Windows accent
        private static readonly Color ColorEmerald = Color.FromArgb(16, 124, 16);     // Windows success
        private static readonly Color ColorAlarm = Color.FromArgb(196, 43, 28);        // Windows danger
        private static readonly Color ColorWarning = Color.FromArgb(245, 158, 11);
        private static readonly Color ColorInput = Color.FromArgb(17, 24, 39);
        private static readonly Font FontBase = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point);
        private static readonly Font FontBold = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
        private static readonly Font FontTitle = new Font("Segoe UI", 11.5F, FontStyle.Bold, GraphicsUnit.Point);

        // Config file stored next to exe (portable: copy exe + .cfg together)
        private static readonly string ConfigFilePath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
            "pm+helper.cfg");

        // Layout DataGridView Column Widths

        // Mock Database Models for User & CardPay
        internal class MockUser
        {
            public string UserId { get; set; }
            public string UserNm { get; set; }
            public string UserPwd { get; set; }
            public string DeptCd { get; set; }
            public string LicNo { get; set; }
        }

        internal class MockCardPay
        {
            public decimal SlipSeq { get; set; }
            public string RecpDt { get; set; }
            public string ChrtNo { get; set; }
            public string CardCoNm { get; set; }
            public decimal CardAmt { get; set; }
            public string CardAdmNo { get; set; }
            public string CardNo { get; set; }
        }

        internal readonly List<MockUser> _mockUserList = new List<MockUser>
        {
            new MockUser { UserId = "admin", UserNm = "관리자", UserPwd = "D033E22AE348AEB5660FC2140AEC35850C4DA997EDC625B5553C99EC2F20D3E90F11D6C59E3F61B78B115F8A22C1A4D4DEC9A887E8F015F07C6418B51E2CD9D2", DeptCd = "SYS", LicNo = "99999" },
            new MockUser { UserId = "doctor1", UserNm = "홍길동", UserPwd = "C6CC809E9F007F10E0557A843513EC2719D3EC5AA9A0A8DA31B4D7BF1DE2871A378EDD13C863B8B82B565D3EF16B78B82F56D2EF8395B9B3222216C1B4E8588A", DeptCd = "MED", LicNo = "12345" },
            new MockUser { UserId = "nurse1", UserNm = "이순신", UserPwd = "FA7A3362145DC9A3D3430A8D9F2E7DCE4AA9AA7AC895C9A31E2A7CE9FA2E87CE9E64F8A3EE7C9A982B565D3EF16B78B82F56D2EF8395B9B3222216C1B4E8588A", DeptCd = "NUR", LicNo = "" }
        };

        internal readonly List<MockCardPay> _mockCardPayList = new List<MockCardPay>
        {
            new MockCardPay { SlipSeq = 10001, RecpDt = "2026-06-15", ChrtNo = "0100028355", CardCoNm = "국민카드", CardAmt = 15000, CardAdmNo = "12345678", CardNo = "4579-xxxx-xxxx-xxxx" },
            new MockCardPay { SlipSeq = 10002, RecpDt = "2026-06-16", ChrtNo = "0000184791", CardCoNm = "신한카드", CardAmt = 8500, CardAdmNo = "87654321", CardNo = "9401-xxxx-xxxx-xxxx" },
            new MockCardPay { SlipSeq = 10003, RecpDt = "2026-06-16", ChrtNo = "0000138658", CardCoNm = "삼성카드", CardAmt = 23000, CardAdmNo = "34567890", CardNo = "5100-xxxx-xxxx-xxxx" }
        };

        // Mock Database Models
        internal class MockDoctor
        {
            public int Seq { get; set; }
            public string Ykiho { get; set; }
            public string YoyangNm { get; set; }
            public string DcId { get; set; }
            public string DcName { get; set; }
            public string DrGubun { get; set; }
        }

        internal readonly List<MockDoctor> _mockDoctorList = new List<MockDoctor>
        {
            new MockDoctor { Seq = 100, Ykiho = "11100079", YoyangNm = "테스트약국", DcId = "101334", DcName = "김의사", DrGubun = "일반의" },
            new MockDoctor { Seq = 101, Ykiho = "11100079", YoyangNm = "테스트약국", DcId = "101334", DcName = "김의사(중복)", DrGubun = "일반의" },
            new MockDoctor { Seq = 102, Ykiho = "11100079", YoyangNm = "테스트약국", DcId = "101334", DcName = "김의사(중복2)", DrGubun = "일반의" },
            new MockDoctor { Seq = 200, Ykiho = "22200088", YoyangNm = "서울병원", DcId = "204488", DcName = "박의사", DrGubun = "전문의" },
            new MockDoctor { Seq = 300, Ykiho = "33300099", YoyangNm = "부산병원", DcId = "305599", DcName = "이의사", DrGubun = "한의사" }
        };

        // Mock Database Models for LabelInfo (TBSIM040_43) & Prescriptions (TBSID040_03)
        internal class MockLabelInfo
        {
            public string DrugCode { get; set; }
            public string Drug { get; set; }
            public string Dan { get; set; }
            public string Save { get; set; }
            public string PrintOp { get; set; }
            public string InputOp { get; set; }
            public string Effct { get; set; }
            public string Comment { get; set; }
            public string SampleUp { get; set; }
            public string EffctUnit { get; set; }
        }

        internal class MockPrescription
        {
            public string DrugSeq { get; set; }
            public string PresDtime { get; set; }
            public string PatNm { get; set; }
            public string PatJuminNo { get; set; }
            public string SunabDt { get; set; }
        }

        internal readonly List<MockLabelInfo> _mockLabelInfoList = new List<MockLabelInfo>
        {
            new MockLabelInfo { DrugCode = "8806446011701", Drug = "타이레놀정500mg", Dan = "정", Save = "실온보관", PrintOp = "1", InputOp = "1", Effct = "해열진통제", Comment = "식후 30분 복용", SampleUp = "0", EffctUnit = "정" },
            new MockLabelInfo { DrugCode = "8806418002409", Drug = "아스피린정100mg", Dan = "정", Save = "실온보관", PrintOp = "1", InputOp = "0", Effct = "소염진통제", Comment = "충분한 물과 함께 복용", SampleUp = "0", EffctUnit = "정" }
        };

        internal readonly List<MockPrescription> _mockPrescriptionList = new List<MockPrescription>
        {
            new MockPrescription { DrugSeq = "PRES202606160001", PresDtime = "2026-06-16 10:30:00", PatNm = "김처방", PatJuminNo = "900101-1234567", SunabDt = "2026-06-16" },
            new MockPrescription { DrugSeq = "PRES202606160002", PresDtime = "2026-06-16 11:15:22", PatNm = "이아픔", PatJuminNo = "850212-2345678", SunabDt = "2026-06-16" },
            new MockPrescription { DrugSeq = "PRES202606160003", PresDtime = "2026-06-16 14:05:40", PatNm = "김처방", PatJuminNo = "900101-1234567", SunabDt = "2026-06-16" }
        };
        internal class MockInventoryItem
        {
            public string DrugCode { get; set; }
            public string DrugName { get; set; }
            public string Manufacturer { get; set; }
            public string Barcode { get; set; }
            public decimal ProperStock { get; set; }
            public decimal TotalStock { get; set; }
            public decimal UnitPrice { get; set; }
        }

        internal readonly List<MockInventoryItem> _mockInventoryList = new List<MockInventoryItem>
        {
            new MockInventoryItem { DrugCode = "ZP00000003", DrugName = "", Manufacturer = "0", Barcode = "8809004779903", ProperStock = 0, TotalStock = 166, UnitPrice = 2240 },
            new MockInventoryItem { DrugCode = "ZP00000029", DrugName = "", Manufacturer = "1", Barcode = "8806113706554", ProperStock = 0, TotalStock = 11, UnitPrice = 3490 },
            new MockInventoryItem { DrugCode = "ZP00000045", DrugName = "", Manufacturer = "1", Barcode = "8806573019812", ProperStock = 0, TotalStock = 3, UnitPrice = 3420 },
            new MockInventoryItem { DrugCode = "644900310", DrugName = "타이레놀정500밀리그램", Manufacturer = "한국존슨앤드존슨", Barcode = "8806449003105", ProperStock = 100, TotalStock = -17604, UnitPrice = 103 },
            new MockInventoryItem { DrugCode = "644900321", DrugName = "아스피린정100밀리그램", Manufacturer = "바이엘코리아", Barcode = "8806449003211", ProperStock = 50, TotalStock = 5500, UnitPrice = 103 },
            new MockInventoryItem { DrugCode = "658600010", DrugName = "아달라트오로스정30", Manufacturer = "바이엘코리아", Barcode = "8806586000104", ProperStock = 20, TotalStock = 4, UnitPrice = 103 },
            new MockInventoryItem { DrugCode = "644900318", DrugName = "신풍아테놀롤정50밀리그램", Manufacturer = "신풍제약", Barcode = "8806449003181", ProperStock = 30, TotalStock = 1300, UnitPrice = 103 },
            new MockInventoryItem { DrugCode = "644900319", DrugName = "신풍아테놀롤정25밀리그램", Manufacturer = "신풍제약", Barcode = "8806449003198", ProperStock = 30, TotalStock = 1600, UnitPrice = 103 }
        };

        // Mock Database Models
        internal class MockRx
        {
            public string ChrtNo { get; set; }
            public string PatNm { get; set; }
            public string PatJuminNo { get; set; }
            public string MedYmd { get; set; }
            public string Medicine { get; set; }
            public string JuminEncrypt { get; set; }
        }

        internal class MockCust
        {
            public string ChrtNo { get; set; }
            public string PatNm { get; set; }
            public string PatJuminNo { get; set; }
            public string Phone { get; set; }
            public string Address { get; set; }
            public string FirstVisit { get; set; }
            public string JuminEncrypt { get; set; }
            public string CusAct { get; set; }
            public string JuminNo { get; set; }
            public string FamNm { get; set; }
            public string HFrDt { get; set; }
            public string HToDt { get; set; }
            public string InsNumber { get; set; }
            public int PatSeq { get; set; }
        }

        private class AggregatedResult
        {
            public string Name { get; set; }
            public string Jumin { get; set; }
            public int Count { get; set; }
            public string LastDate { get; set; }
        }

        // Mock Database Lists
        internal readonly List<MockRx> _mockRxList = new List<MockRx>
        {
            // 김승학 과거이력 데모 데이터
            new MockRx { ChrtNo = "0000999999", PatNm = "김승학", PatJuminNo = "650820-1******", MedYmd = "2026-06-17", Medicine = "아스피린정 100mg", JuminEncrypt = "ENC_KIM_SH" },

            // Case 1: Chart 0000184791 (Cheon Mi-seon & Park Bok-soon)
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2025-08-01", Medicine = "리피토정 10mg, 아토젯정 10/10mg", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2025-08-22", Medicine = "리피토정 10mg, 크레스토정 5mg", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2025-10-24", Medicine = "리피토정 10mg, 노바스크정 5mg", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2025-11-28", Medicine = "리피토정 10mg, 가스모틴정 5mg", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2026-01-02", Medicine = "리피토정 10mg, 엘도스캡슐", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2026-03-06", Medicine = "리피토정 10mg, 뮤코펙트정", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "천미선", PatJuminNo = "770315-2******", MedYmd = "2026-05-11", Medicine = "리피토정 10mg, 아스피린정 100mg", JuminEncrypt = "koLqZr1Kx3+gYiBY8G6VKYm7KoHzqk6cYYPTZgAk7Hc=01" },
            new MockRx { ChrtNo = "0000184791", PatNm = "박복순", PatJuminNo = "590307-2******", MedYmd = "2026-06-05", Medicine = "타이레놀정 500mg", JuminEncrypt = "b5VH0jLuaoKjK1pCcm4D0BSp1ywytRCGJK92vmG1ym0=01" },

            // Case 2: Chart 0000138658 (Kim Hyun-sook, 2 jumins)
            new MockRx { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "360115-2******", MedYmd = "2026-05-01", Medicine = "크레스토정 10mg", JuminEncrypt = "ENC_KIM_36" },
            new MockRx { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "360115-2******", MedYmd = "2026-05-15", Medicine = "크레스토정 10mg", JuminEncrypt = "ENC_KIM_36" },
            new MockRx { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "360115-2******", MedYmd = "2026-06-01", Medicine = "크레스토정 10mg", JuminEncrypt = "ENC_KIM_36" },
            new MockRx { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "360115-2******", MedYmd = "2026-06-02", Medicine = "크레스토정 10mg", JuminEncrypt = "ENC_KIM_36" },
            new MockRx { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "411129-2******", MedYmd = "2026-05-10", Medicine = "아스피린정 100mg", JuminEncrypt = "ENC_KIM_41" },

            // Case 3: Chart 0000187987 (Same Jumin, 2 names - Choo Nam-young & Choo Moo-gyeom)
            new MockRx { ChrtNo = "0000187987", PatNm = "추남영", PatJuminNo = "450815-1******", MedYmd = "2026-04-10", Medicine = "디옥타정, 뮤코펙트정", JuminEncrypt = "ENC_CHOO" },
            new MockRx { ChrtNo = "0000187987", PatNm = "추무겸", PatJuminNo = "450815-1******", MedYmd = "2026-05-20", Medicine = "가스모틴정 5mg", JuminEncrypt = "ENC_CHOO" },

            // Case 4: Chart 0000208778 (Ahn Jong-hak & Jang Ji-na)
            new MockRx { ChrtNo = "0000208778", PatNm = "안종학", PatJuminNo = "520412-1******", MedYmd = "2026-03-01", Medicine = "노바스크정 5mg", JuminEncrypt = "ENC_AHN" },
            new MockRx { ChrtNo = "0000208778", PatNm = "장지나", PatJuminNo = "850630-2******", MedYmd = "2026-04-05", Medicine = "타이레놀정 500mg", JuminEncrypt = "ENC_JANG" },

            // Standard Patients
            new MockRx { ChrtNo = "0100028355", PatNm = "박순영", PatJuminNo = "571029-2345678", MedYmd = "2026-06-01", Medicine = "아스피린정 100mg, 코대원포르테시럽", JuminEncrypt = "ENC_5710292345678" }
        };

        internal readonly List<MockCust> _mockCustList = new List<MockCust>
        {
            // Case 1 Customers
            new MockCust { ChrtNo = "0000184791", PatNm = "박복순", PatJuminNo = "590307-2******", Phone = "010-9999-8888", Address = "대전광역시 서구 둔산동 100", FirstVisit = "2025-08-01", JuminEncrypt = "b5VH0jLuaoKjK1pCcm4D0BSp1ywytRCGJK92vmG1ym0=01", CusAct = "1", JuminNo = "5903072******" },
            new MockCust { ChrtNo = "0000144177", PatNm = "박복순", PatJuminNo = "590307-2******", Phone = "010-9999-7777", Address = "대전광역시 서구 탄방동 200", FirstVisit = "2023-04-10", JuminEncrypt = "b5VH0jLuaoKjK1pCcm4D0BSp1ywytRCGJK92vmG1ym0=01", CusAct = "1", JuminNo = "5903072******" },

            // Case 2 Customers
            new MockCust { ChrtNo = "0000138658", PatNm = "김현숙", PatJuminNo = "411129-2******", Phone = "010-1111-2222", Address = "대전광역시 중구 은행동 10", FirstVisit = "2024-02-15", JuminEncrypt = "ENC_KIM_41", CusAct = "1", JuminNo = "4111292******" },
            new MockCust { ChrtNo = "0000138668", PatNm = "김현숙", PatJuminNo = "621215-2******", Phone = "010-1111-3333", Address = "대전광역시 중구 선화동 20", FirstVisit = "2024-06-20", JuminEncrypt = "ENC_KIM_62", CusAct = "1", JuminNo = "6212152******" },

            new MockCust { ChrtNo = "0100028355", PatNm = "박순영", PatJuminNo = "571029-2345678", Phone = "010-1234-5678", Address = "서울특별시 강남구 역삼동 123-45", FirstVisit = "2020-03-12", JuminEncrypt = "ENC_5710292345678", CusAct = "1", JuminNo = "571029-2345678" },
            new MockCust { ChrtNo = "0000999999", PatNm = "김승학", PatJuminNo = "650820-1******", Phone = "010-3333-4444", Address = "부산광역시 해운대구 우동 20", FirstVisit = "2021-07-20", JuminEncrypt = "ENC_KIM_SH", CusAct = "0", JuminNo = "6508201******", FamNm = "김승학", HFrDt = "2021-08-20", HToDt = "2021-07-20", InsNumber = "81116283481", PatSeq = 3 },
            new MockCust { ChrtNo = "0100026131", PatNm = "", PatJuminNo = "******-******", Phone = "", Address = "", FirstVisit = "2022-09-12", JuminEncrypt = "kK9LhrP2HFOC+IcfVnNMTg==01", CusAct = "1", JuminNo = "******-******" },
            new MockCust { ChrtNo = "0100026132", PatNm = "", PatJuminNo = "******-******", Phone = "", Address = "", FirstVisit = "2022-09-12", JuminEncrypt = "kK9LhrP2HFOC+IcfVnNMTg==01", CusAct = "1", JuminNo = "******-******" },
            new MockCust { ChrtNo = "0100026133", PatNm = "", PatJuminNo = "******-******", Phone = "", Address = "", FirstVisit = "2022-09-12", JuminEncrypt = "kK9LhrP2HFOC+IcfVnNMTg==01", CusAct = "1", JuminNo = "******-******" }
        };

        // UI Controls - Settings Area
        private CheckBox _chkDemoMode;
        private TextBox _txtServer;
        private CheckBox _chkIntegratedSecurity;
        private TextBox _txtUser;
        private TextBox _txtPassword;
        private ComboBox _cmbDatabases;
        private Button _btnLoadDbs;
        private Button _btnSaveConfig;
        private Panel _pnlCredentials;
        private Label _lblStatusBadge;
        private Button _btnMyBox;
        private Label _lblVersionBadge;
        private Button _btnCheckUpdate;
        private Label _lblUpdateBadge;
        private ToolTip _toolTip;
        private MenuStrip _menuStrip;
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _statusMainLabel;

        // UI Controls - Step 1 (Search Prescription)
        private TextBox _txtSearchName;
        private TextBox _txtSearchJumin;
        private Button _btnSearchRx;
        private DataGridView _dgvResults;

        // UI Controls - Step 2 (Verify Customer)
        private TextBox _txtSearchChrtNo;
        private Button _btnSearchCust;
        private Button _btnRestoreCust;
        private TabControl _tabControl;
        private TabPage _tabMainWorkspace;
        private TabPage _tabTroubleshooter;
        private TabPage _tabDoctorLicense;
        private TabPage _tabDbRecovery;
        private CheckBox _chkDbPmMain;
        private CheckBox _chkDbPmplusDums;
        private CheckBox _chkDbPmplusImage;
        private CheckBox _chkDbPmplusJoblog;
        private Button _btnRunDbRecovery;
        private Button _btnShrinkDb;
        private Button _btnShrinkLog;
        private Button _btnDropDrugUpdateDb;
        private Button _btnDropDurUpdateDb;
        private TextBox _txtDbRecoveryLog;
        private DataGridView _dgvDoctors;
        private Button _btnSearchDoctors;
        private Button _btnDeleteDoctors;
        private Label _lblDoctorStatus;
        private TroubleshooterForm _troubleshooter;
        private Label _lblCustChrtNo;
        private Label _lblCustName;
        private Label _lblCustNameTitle;
        private Label _lblCustChrtNoBadge;
        private Label _lblCustJumin;
        private Label _lblCustPhone;
        private Label _lblCustAddress;
        private Label _lblCustFirstVisit;
        private ListBox _lstRxHistory;
        private Label _lblToast;
        private Timer _toastTimer;
        private string _detectedDateColumn = null;

        // UI Controls - Narcotics Management
        private TabPage _tabNarcoticsManagement;
        private TabControl _subTabNarcotics;
        private TabPage _tabSeqCorrection;
        private DataGridView _dgvNarcoticErrors;
        private Button _btnScanNarcotics;
        private Button _btnFixSelectedNarcotic;
        private Button _btnFixAllNarcotics;
        private TextBox _txtNarcoticsLog;

        // UI Controls - Narcotics Usage Quantity Cleanup
        private TabPage _tabUsageQuantity;
        private Button _btnSearchGhostDates;
        private Button _btnDeleteGhostRecords;
        private TextBox _txtGhostDatesResult;
        private Button _btnScanCanceledPrescs;
        private Button _btnDeleteSelectedCanceled;
        private Button _btnDeleteAllCanceled;
        private DataGridView _dgvCanceledPrescs;
        private TextBox _txtUsageQuantityLog;

        // UI Controls - SQL Query Runner (Top-level Tab)
        private TabPage _tabQueryRunner;
        private ComboBox _cmbQueryDbSelector;
        private TextBox _txtQueryInput;
        private Button _btnExecuteQuery;
        private DataGridView _dgvQueryResult;
        private Label _lblQueryStatus;

        // UI Controls - Data Management (CRUD)
        private TabPage _tabDataManagement;
        private TabControl _subTabDataManagement;
        private DataGridView _dgvUsers;
        private TextBox _txtUserSearchId;
        private TextBox _txtUserSearchName;
        private TextBox _txtUserId;
        private TextBox _txtUserNm;
        private TextBox _txtUserPwd;
        private TextBox _txtUserDeptCd;
        private TextBox _txtUserLicNo;
        private Button _btnUserSearch;
        private Button _btnUserAdd;
        private Button _btnUserUpdate;
        private Button _btnUserDelete;
        private Button _btnUserClear;

        private DataGridView _dgvCardPays;
        private TextBox _txtCardSearchChart;
        private TextBox _txtCardSearchDate;
        private TextBox _txtCardSlipSeq;
        private TextBox _txtCardRecpDt;
        private TextBox _txtCardChrtNo;
        private TextBox _txtCardCoNm;
        private TextBox _txtCardAmt;
        private TextBox _txtCardAdmNo;
        private TextBox _txtCardNo;
        private Button _btnCardSearch;
        private Button _btnCardAdd;
        private Button _btnCardUpdate;
        private Button _btnCardDelete;
        private Button _btnCardClear;

        // UI Controls - LabelInfo (TBSIM040_43)
        private DataGridView _dgvLabelInfos;
        private TextBox _txtLabelSearchCode;
        private TextBox _txtLabelSearchName;
        private TextBox _txtLabelDrugCode;
        private TextBox _txtLabelDrug;
        private TextBox _txtLabelDan;
        private TextBox _txtLabelSave;
        private TextBox _txtLabelPrintOp;
        private TextBox _txtLabelInputOp;
        private TextBox _txtLabelEffct;
        private TextBox _txtLabelComment;
        private TextBox _txtLabelSampleUp;
        private TextBox _txtLabelEffctUnit;
        private Button _btnLabelSearch;
        private Button _btnLabelAdd;
        private Button _btnLabelUpdate;
        private Button _btnLabelDelete;
        private Button _btnLabelClear;

        // UI Controls - Inventory Status
        private TabPage _tabInventoryManagement;
        private TabControl _subTabInventoryManagement;
        private DataGridView _dgvStockMovementErrors;
        private DataGridView _dgvStockAuditDrugSearch;
        private TextBox _txtStockAuditDrugName;
        private TextBox _txtStockAuditDrugCode;
        private TextBox _txtStockAuditUnit;
        private TextBox _txtStockAuditMinQty;
        private TextBox _txtStockAuditDrugInfo;
        private Button _btnStockAuditDrugSearch;
        private Button _btnStockAuditRun;
        private DataGridView _dgvInventory;
        private TextBox _txtInventorySearch;
        private CheckBox _chkInventoryNoNameOnly;
        private CheckBox _chkInventoryExcludeZeroStock;
        private Button _btnInventorySearch;
        private TextBox _txtInvFormDrugCode;
        private TextBox _txtInvFormBarcode;
        private TextBox _txtInvFormDrugName;
        private TextBox _txtInvFormManufacturer;
        private Button _btnInvFormUpdate;
        private Button _btnInvFormDelete;
        private Button _btnInvBatchDelete;
        private Button _btnInvCleanDupBarcodes;
        private Button _btnDurakanAudit;
        private Button _btnInvBarcodeSearchWeb;
        private Label _lblInvFormSuggest;
        private SplitContainer _splitInventory;

        // UI Controls - Stock Adjustment Restore
        private TabPage _tabStockAdjustmentRestore;
        private Label _lblStockAdjBackupStatus;
        private Button _btnStockAdjAttachBackup;
        private Button _btnStockAdjDetachBackup;
        private Button _btnStockAdjScan;
        private Button _btnStockAdjRestoreSelected;
        private Button _btnStockAdjRestoreAll;
        private Button _btnStockAdjExportCsv;
        private DataGridView _dgvStockAdjSummary;
        private DataGridView _dgvStockAdjDetail;
        private Label _lblStockAdjDetailTitle;
        private Label _lblStockAdjSummaryCount;
        private Button _btnStockAdjSelectAll;
        private Button _btnStockAdjDeselectAll;
        private Button _btnStockAdjSelectMissingOnly;
        private SplitContainer _splitStockAdj;
        private int _distStockAdj = 520;
        private DataTable _stockAdjSummaryDt;

        // UI Controls - Prescription Delete (TBSID040_03, 04, 05)
        private TabPage _tabPrescriptionDelete;
        private TextBox _txtRxDelSearchName;
        private TextBox _txtRxDelSearchJumin;
        private Button _btnRxDelSearch;
        private DataGridView _dgvRxDeleteList;
        private Button _btnRxDeleteExecute;

        // UI Controls - Monthly dispensing/claim comparison
        private TabPage _tabClaimComparison;
        private DateTimePicker _dtpClaimComparisonMonth;
        private ComboBox _cmbClaimComparisonType;
        private CheckBox _chkClaimComparisonExcludeZero;
        private Button _btnClaimComparisonSearch;
        private Button _btnClaimComparisonExport;
        private DataGridView _dgvClaimComparison;
        private Label _lblClaimComparisonSummary;

        // UI Controls - Log / Claim Mismatch Scanner
        private TabPage _tabLogClaimMismatch;
        private ComboBox _cmbLogMismatchFilter;
        private TextBox _txtLogMismatchTarget;
        private Button _btnLogMismatchScan;
        private Button _btnLogMismatchExport;
        private Button _btnAttachPrescriptionBackup;
        private Button _btnDetachPrescriptionBackup;
        private Label _lblBackupConnectionStatus;
        private DataGridView _dgvLogMismatchSummary;
        private Label _lblLogMismatchSummary;
        private Label _lblLogMismatchDetailInfo;
        private DataGridView _dgvLogMismatchDetail;
        private TextBox _txtLogRestoreNewChrtNo;
        private Button _btnLogRestoreSeparate;
        private Button _btnLogRestoreSelectAll;
        private Button _btnLogRestoreDeselectAll;
        private ComboBox _cmbLogRestorePatientGroup;
        private Panel _pnlLogMismatchDetailAction;
        private FlowLayoutPanel _pnlJuminClassificationViews;
        private Button _btnJuminShowRestoreTargets;
        private Button _btnJuminShowNoEvidence;
        private Button _btnJuminShowUnidentified;
        private Label _lblJuminNormalCount;
        private DataTable _juminClassificationAll;
        private string _juminClassificationView = "복구 대상";
        private Label _lblLogRestorePatient;
        private Label _lblLogRestoreNewChart;
        private SplitContainer _splitLogMismatch;
        private int _distLogMismatch = 480;

        // UI Controls - Dispense Customer Management
        private TabPage _tabDispenseCustomerManagement;
        private TabControl _subTabDispenseCustomer;
        internal TabPage _tabPastHistoryManagement;
        public TextBox _txtHistoryChartNo;
        private bool _isDemo;
        private DataGridView _dgvHistoryMaster;
        private Button _btnHistorySearch;
        private Button _btnHistorySave;
        private Button _btnHistoryDelete;

        // UI Controls - SQL Service Control
        private Label _lblSqlServiceStatus;
        private Button _btnSqlServiceStart;
        private Button _btnSqlServiceStop;
        private Timer _sqlServiceTimer;
        private string _lastSqlServiceStatus = "UNKNOWN"; // For demo mode mock behavior

        // UI Controls - SplitContainers for layouts and their distance settings
        internal SplitContainer _splitChartResolver;
        private SplitContainer _splitUser;
        private SplitContainer _splitCard;
        private SplitContainer _splitLabel;
        private SplitContainer _splitRx;

        internal int _distChartResolver = 460;
        private int _distUser = 550;
        private int _distCard = 550;
        private int _distLabel = 550;
        private int _distRx = 550;

        public MainForm()
        {
            InitializeComponent();
            this.FormClosing += MainForm_FormClosing;
            try
            {
                this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { }
            LoadConfig();
            RefreshAttachedBackupStatus();

            // Instantiate embedded Troubleshooter Form
            _troubleshooter = new TroubleshooterForm(this, _chkDemoMode.Checked)
            {
                TopLevel = false,
                FormBorderStyle = FormBorderStyle.None,
                Dock = DockStyle.Fill
            };
            _tabTroubleshooter.Controls.Add(_troubleshooter);
            _troubleshooter.Show();

            ApplyModernUiEnhancements();
            ToggleDemoMode(_chkDemoMode.Checked);

            // SQL Service Monitor Timer
            _sqlServiceTimer = new Timer();
            _sqlServiceTimer.Interval = 3000;
            _sqlServiceTimer.Tick += (s, e) => UpdateSqlServiceUI();
            _sqlServiceTimer.Start();

            UpdateSqlServiceUI();
        }

        private void ApplyModernUiEnhancements()
        {
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.MinimumSize = new Size(1120, 720);
            this.Font = FontBase;
            this.Text = "pm+helper v" + UpdateManager.CurrentVersion + " - 환자 차트/DB 유지보수 도우미";

            _toolTip = new ToolTip
            {
                AutoPopDelay = 12000,
                InitialDelay = 350,
                ReshowDelay = 100,
                ShowAlways = true
            };

            ApplyModernStyleRecursive(this);
            LayoutLogMismatchActionControls();
            WireTabDrawHandlers(this);
            AddBeginnerHelpStrip();
            ConfigureImportantTooltips();
            UpdateStatus("준비됨 - 먼저 데모 모드에서 흐름을 확인한 뒤 실제 DB에 연결하세요.");
        }

        private void AddModernMenuStrip()
        {
            if (_menuStrip != null) return;

            _menuStrip = new MenuStrip
            {
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                Font = FontBase,
                Padding = new Padding(8, 4, 8, 4),
                GripStyle = ToolStripGripStyle.Hidden,
                RenderMode = ToolStripRenderMode.Professional,
                Renderer = new ModernMenuRenderer()
            };

            ToolStripMenuItem fileMenu = new ToolStripMenuItem("파일");
            ToolStripMenuItem saveItem = new ToolStripMenuItem("설정 저장", CreateMenuIcon(ColorIndigo));
            saveItem.Click += (s, e) => _btnSaveConfig.PerformClick();
            ToolStripMenuItem exitItem = new ToolStripMenuItem("종료");
            exitItem.Click += (s, e) => this.Close();
            fileMenu.DropDownItems.Add(saveItem);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(exitItem);

            ToolStripMenuItem dbMenu = new ToolStripMenuItem("DB");
            ToolStripMenuItem loadDbItem = new ToolStripMenuItem("DB 목록 불러오기", CreateMenuIcon(ColorEmerald));
            loadDbItem.Click += (s, e) => _btnLoadDbs.PerformClick();
            dbMenu.DropDownItems.Add(loadDbItem);

            ToolStripMenuItem viewMenu = new ToolStripMenuItem("보기");
            ToolStripMenuItem demoItem = new ToolStripMenuItem("데모 모드 전환", CreateMenuIcon(ColorWarning));
            demoItem.Click += (s, e) => _chkDemoMode.Checked = !_chkDemoMode.Checked;
            viewMenu.DropDownItems.Add(demoItem);

            _menuStrip.Items.Add(fileMenu);
            _menuStrip.Items.Add(dbMenu);
            _menuStrip.Items.Add(viewMenu);
            this.MainMenuStrip = _menuStrip;
            this.Controls.Add(_menuStrip);
            _menuStrip.BringToFront();
        }

        private void AddBeginnerHelpStrip()
        {
            if (_statusStrip != null) return;

            _statusStrip = new StatusStrip
            {
                BackColor = ColorBgCard,
                ForeColor = ColorTextSec,
                SizingGrip = false
            };
            _statusMainLabel = new ToolStripStatusLabel
            {
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = ColorTextSec
            };
            ToolStripStatusLabel helpLabel = new ToolStripStatusLabel
            {
                Text = "작업 순서: 1) 데모 확인  2) DB 연결  3) 조회  4) 대상 검토  5) 복구/삭제 실행",
                ForeColor = ColorTextSec
            };
            _statusStrip.Items.Add(_statusMainLabel);
            _statusStrip.Items.Add(helpLabel);
            this.Controls.Add(_statusStrip);
            _statusStrip.BringToFront();
        }

        private void ConfigureImportantTooltips()
        {
            SetTip(_chkDemoMode, "실제 DB를 건드리지 않고 기능 흐름을 연습합니다. 초보자는 먼저 이 모드로 사용하세요.");
            SetTip(_txtServer, "SQL Server 주소입니다. 예: .\\pmplus20");
            SetTip(_chkIntegratedSecurity, "Windows 로그인 계정으로 SQL Server에 접속합니다.");
            SetTip(_btnLoadDbs, "서버에 연결해 데이터베이스 목록을 불러옵니다.");
            SetTip(_cmbDatabases, "작업 대상 데이터베이스입니다. 잘못 선택하면 다른 DB에 작업될 수 있습니다.");
            SetTip(_btnSaveConfig, "현재 연결 정보와 화면 분할 위치를 저장합니다.");
            SetTip(_btnSearchRx, "환자명 또는 주민번호 앞 7자리로 처방 이력을 검색합니다.");
            SetTip(_btnSearchCust, "입력한 차트번호의 고객 마스터 정보를 조회합니다.");
            SetTip(_btnRestoreCust, "선택한 처방 정보를 기준으로 고객 마스터 정보를 복구합니다. 실행 전 내용을 반드시 확인하세요.");
            SetTip(_btnRunDbRecovery, "DB 응급 복구 작업입니다. 백업이 없으면 실행하지 마세요.");
            SetTip(_btnShrinkDb, "DB 파일 크기를 줄입니다. 업무 시간 외 실행을 권장합니다.");
            SetTip(_btnShrinkLog, "로그 파일 크기를 줄입니다. 백업 정책을 먼저 확인하세요.");
            SetTip(_btnExecuteQuery, "SQL을 실행합니다. UPDATE/DELETE/DROP 등 변경 쿼리는 매우 주의하세요.");
        }

        private void SetTip(Control control, string text)
        {
            if (_toolTip != null && control != null) _toolTip.SetToolTip(control, text);
        }

        private void ApplyModernStyleRecursive(Control root)
        {
            foreach (Control child in root.Controls)
            {
                ApplyModernStyle(child);
                if (child.HasChildren) ApplyModernStyleRecursive(child);
            }
        }

        private void ApplyModernStyle(Control control)
        {
            if (control is Button)
            {
                Button btn = (Button)control;
                btn.Cursor = Cursors.Hand;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.Font = FontBold;
                if (btn.Height < 30) btn.Height = 30;
                btn.Padding = new Padding(10, 0, 10, 0);
                btn.Margin = new Padding(4);
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                btn.Image = null;
                FitButtonToContent(btn);

                string text = btn.Text ?? "";
                if (text.Contains("삭제") || text.Contains("중지") || text.Contains("응급"))
                {
                    btn.BackColor = ColorAlarm;
                    btn.ForeColor = Color.White;
                }
                else if (text.Contains("복구") || text.Contains("저장") || text.Contains("조회") || text.Contains("검색") || text.Contains("실행"))
                {
                    btn.BackColor = ColorIndigo;
                    btn.ForeColor = Color.White;
                }
            }
            else if (control is TextBox)
            {
                TextBox txt = (TextBox)control;
                txt.BackColor = txt.ReadOnly ? Color.FromArgb(12, 18, 29) : ColorInput;
                txt.ForeColor = txt.ReadOnly ? ColorTextSec : ColorTextMain;
                txt.BorderStyle = BorderStyle.FixedSingle;
                txt.Font = txt.Multiline ? new Font("Consolas", 9.5F) : FontBase;
                txt.Margin = new Padding(4);
            }
            else if (control is ComboBox)
            {
                ComboBox cmb = (ComboBox)control;
                cmb.BackColor = ColorInput;
                cmb.ForeColor = ColorTextMain;
                cmb.FlatStyle = FlatStyle.Flat;
                cmb.Font = FontBase;
                cmb.Margin = new Padding(4);
            }
            else if (control is CheckBox)
            {
                CheckBox chk = (CheckBox)control;
                chk.ForeColor = ColorTextMain;
                chk.Font = FontBase;
                chk.Margin = new Padding(4);
            }
            else if (control is Label)
            {
                Label lbl = (Label)control;
                if (lbl.ForeColor == SystemColors.ControlText) lbl.ForeColor = ColorTextMain;
                if (lbl.Font == null || lbl.Font.Name != "Segoe UI") lbl.Font = FontBase;
                lbl.Margin = new Padding(4);
            }
            else if (control is DataGridView)
            {
                ApplyGridStyle((DataGridView)control);
            }
            else if (control is GroupBox)
            {
                ApplyCardGroupBox((GroupBox)control);
            }
            else if (control is TabPage)
            {
                control.BackColor = ColorBgMain;
            }
            else if (control is Panel || control is SplitterPanel)
            {
                if (control.BackColor == SystemColors.Control) control.BackColor = ColorBgMain;
            }
        }

        private void FitButtonToContent(Button btn)
        {
            int minWidth = GetButtonContentWidth(btn);
            if (btn.Width < minWidth)
            {
                btn.Width = minWidth;
            }
        }

        private int GetButtonContentWidth(Button btn)
        {
            string text = btn.Text ?? "";
            Font measureFont = btn.Font != null && btn.Font.Style == FontStyle.Bold ? btn.Font : FontBold;
            Size textSize = TextRenderer.MeasureText(text, measureFont);
            bool isTabNavButton = btn.Tag != null && btn.Tag.ToString() == "TabNav";
            int imageWidth = 0;
            int horizontalPadding = Math.Max(btn.Padding.Left + btn.Padding.Right, 20);
            return textSize.Width + imageWidth + horizontalPadding + 16;
        }

        private Image CreateButtonIcon(Button btn)
        {
            string text = btn.Text ?? "";
            Color color = Color.White;
            if (text.Contains("삭제") || text.Contains("중지") || text.Contains("응급")) color = Color.FromArgb(255, 224, 224);
            else if (text.Contains("저장")) color = Color.FromArgb(220, 235, 255);
            else if (text.Contains("조회") || text.Contains("검색")) color = Color.FromArgb(225, 242, 255);
            else if (text.Contains("시작") || text.Contains("복구")) color = Color.FromArgb(224, 255, 232);
            return CreateMenuIcon(color);
        }

        private static Bitmap CreateMenuIcon(Color color)
        {
            Bitmap bmp = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bmp))
            using (SolidBrush brush = new SolidBrush(color))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.FillEllipse(brush, 3, 3, 10, 10);
            }
            return bmp;
        }

        private void ApplyCardGroupBox(GroupBox gb)
        {
            gb.BackColor = ColorBgCard;
            gb.ForeColor = ColorTextMain;
            gb.Font = FontBold;
            gb.Padding = new Padding(14, 18, 14, 14);
            gb.Margin = new Padding(8);
            gb.Paint -= CardGroupBox_Paint;
            gb.Paint += CardGroupBox_Paint;
        }

        private void CardGroupBox_Paint(object sender, PaintEventArgs e)
        {
            GroupBox gb = (GroupBox)sender;
            e.Graphics.Clear(gb.Parent != null ? gb.Parent.BackColor : ColorBgMain);
            Rectangle card = new Rectangle(0, 8, gb.Width - 1, gb.Height - 9);
            using (SolidBrush bg = new SolidBrush(ColorBgCard))
            using (Pen border = new Pen(ColorBorder))
            using (SolidBrush text = new SolidBrush(ColorTextMain))
            {
                e.Graphics.FillRectangle(bg, card);
                e.Graphics.DrawRectangle(border, card);
                e.Graphics.DrawString(gb.Text, FontBold, text, new PointF(12, 0));
            }
        }

        private void ApplyGridStyle(DataGridView dgv)
        {
            dgv.BackgroundColor = ColorBgCard;
            dgv.BorderStyle = BorderStyle.None;
            dgv.GridColor = ColorBorder;
            dgv.RowHeadersVisible = false;
            dgv.EnableHeadersVisualStyles = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgv.RowTemplate.Height = 30;
            dgv.ColumnHeadersHeight = 36;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(20, 29, 43);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(31, 41, 55);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontBold;
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 4, 8, 4);
            dgv.DefaultCellStyle.BackColor = ColorBgCard;
            dgv.DefaultCellStyle.ForeColor = ColorTextMain;
            dgv.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);
        }

        private void WireTabDrawHandlers(Control root)
        {
            foreach (Control child in root.Controls)
            {
                if (child is TabControl)
                {
                    TabControl tc = (TabControl)child;
                    if (tc.Tag != null && tc.Tag.ToString() == "HiddenTabs")
                    {
                        tc.Appearance = TabAppearance.Buttons;
                        tc.DrawMode = TabDrawMode.Normal;
                        tc.SizeMode = TabSizeMode.Fixed;
                        tc.ItemSize = new Size(1, 1);
                        tc.Multiline = false;
                        tc.Padding = new Point(0, 0);
                        if (child.HasChildren) WireTabDrawHandlers(child);
                        continue;
                    }

                    tc.Appearance = TabAppearance.FlatButtons;
                    tc.DrawMode = TabDrawMode.OwnerDrawFixed;
                    tc.SizeMode = TabSizeMode.Fixed;
                    tc.ItemSize = new Size(220, 34);
                    tc.Multiline = true;
                    tc.Font = FontBase;
                    tc.Padding = new Point(12, 6);
                    tc.DrawItem -= ModernTabControl_DrawItem;
                    tc.DrawItem += ModernTabControl_DrawItem;
                }
                if (child.HasChildren) WireTabDrawHandlers(child);
            }
        }

        private void InstallVisibleTabNavigation(Panel host, TabControl tabControl)
        {
            if (host == null || tabControl == null) return;

            FlowLayoutPanel nav = new FlowLayoutPanel
            {
                Dock = DockStyle.None,
                Height = 52,
                BackColor = ColorBgMain,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(10, 8, 10, 8),
                Margin = new Padding(0)
            };

            tabControl.Tag = "HiddenTabs";
            tabControl.Appearance = TabAppearance.Buttons;
            tabControl.SizeMode = TabSizeMode.Fixed;
            tabControl.ItemSize = new Size(1, 1);
            tabControl.Dock = DockStyle.None;

            List<Button> navButtons = new List<Button>();
            for (int i = 0; i < tabControl.TabPages.Count; i++)
            {
                int tabIndex = i;
                string text = tabControl.TabPages[i].Text;
                Size textSize = TextRenderer.MeasureText(text, FontBold);
                Button btn = new Button
                {
                    Text = text,
                    Tag = "TabNav",
                    Width = Math.Max(150, Math.Min(250, textSize.Width + 36)),
                    Height = 36,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = i == tabControl.SelectedIndex ? ColorIndigo : ColorBgCard,
                    ForeColor = i == tabControl.SelectedIndex ? Color.White : ColorTextMain,
                    Font = FontBold,
                    Margin = new Padding(3),
                    TextAlign = ContentAlignment.MiddleCenter,
                    UseVisualStyleBackColor = false
                };
                btn.FlatAppearance.BorderSize = 1;
                btn.FlatAppearance.BorderColor = i == tabControl.SelectedIndex ? Color.FromArgb(96, 165, 250) : ColorBorder;
                btn.Click += delegate { tabControl.SelectedIndex = tabIndex; };
                navButtons.Add(btn);
                nav.Controls.Add(btn);
            }

            tabControl.SelectedIndexChanged += delegate
            {
                for (int i = 0; i < navButtons.Count; i++)
                {
                    bool selected = i == tabControl.SelectedIndex;
                    navButtons[i].BackColor = selected ? ColorIndigo : ColorBgCard;
                    navButtons[i].ForeColor = selected ? Color.White : ColorTextMain;
                    navButtons[i].FlatAppearance.BorderColor = selected ? Color.FromArgb(96, 165, 250) : ColorBorder;
                }
            };

            Action layout = delegate
            {
                int hostWidth = Math.Max(0, host.ClientSize.Width);
                int hostHeight = Math.Max(0, host.ClientSize.Height);
                int menuOffset = host.Parent is Form ? (_menuStrip != null ? _menuStrip.Height : 0) : 0;
                int columns = Math.Max(1, hostWidth / 180);
                int rows = Math.Max(1, (nav.Controls.Count + columns - 1) / columns);
                int navHeight = Math.Min(132, 16 + rows * 42);
                int contentTop = menuOffset + navHeight;

                nav.Bounds = new Rectangle(0, menuOffset, hostWidth, navHeight);
                tabControl.Bounds = new Rectangle(0, contentTop, hostWidth, Math.Max(0, hostHeight - contentTop));
            };

            host.Controls.Add(nav);
            nav.BringToFront();
            host.Resize += delegate { layout(); };
            layout();
        }

        private void ModernTabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl tc = (TabControl)sender;
            TabPage page = tc.TabPages[e.Index];
            bool selected = e.Index == tc.SelectedIndex;
            Rectangle rect = e.Bounds;
            Rectangle bgRect = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height);
            Rectangle textRect = new Rectangle(rect.X + 10, rect.Y + 4, rect.Width - 20, rect.Height - 8);

            using (SolidBrush baseBg = new SolidBrush(ColorBgMain))
            using (SolidBrush bg = new SolidBrush(selected ? ColorIndigo : ColorBgCard))
            using (Pen border = new Pen(selected ? Color.FromArgb(96, 165, 250) : ColorBorder))
            {
                e.Graphics.FillRectangle(baseBg, e.Bounds);
                e.Graphics.FillRectangle(bg, bgRect);
                e.Graphics.DrawRectangle(border, bgRect.X, bgRect.Y, bgRect.Width - 1, bgRect.Height - 1);
                TextRenderer.DrawText(
                    e.Graphics,
                    page.Text,
                    FontBold,
                    textRect,
                    selected ? Color.White : ColorTextMain,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
        }

        private void UpdateStatus(string message)
        {
            if (_statusMainLabel != null)
            {
                _statusMainLabel.Text = message;
            }
        }

        private class ModernMenuRenderer : ToolStripProfessionalRenderer
        {
            public ModernMenuRenderer()
                : base(new ModernMenuColorTable())
            {
            }
        }

        private class ModernMenuColorTable : ProfessionalColorTable
        {
            public override Color MenuItemSelected { get { return ColorIndigo; } }
            public override Color MenuItemBorder { get { return ColorIndigo; } }
            public override Color MenuBorder { get { return ColorBorder; } }
            public override Color ToolStripDropDownBackground { get { return ColorBgCard; } }
            public override Color ImageMarginGradientBegin { get { return ColorBgCard; } }
            public override Color ImageMarginGradientMiddle { get { return ColorBgCard; } }
            public override Color ImageMarginGradientEnd { get { return ColorBgCard; } }
            public override Color SeparatorDark { get { return ColorBorder; } }
            public override Color SeparatorLight { get { return ColorBorder; } }
        }

        private void InitializeComponent()
        {
            // Main Form Settings
            this.Text = "pm+helper v" + UpdateManager.CurrentVersion + " - 환자 차트/DB 유지보수 도우미";
            this.Size = new Size(1280, 768);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = ColorBgMain;
            this.ForeColor = ColorTextMain;
            this.Font = new Font("맑은 고딕", 9.75F, FontStyle.Regular, GraphicsUnit.Point);

            // Toast Timer Setup
            _toastTimer = new Timer();
            _toastTimer.Interval = 3000;
            _toastTimer.Tick += (s, e) => {
                _lblToast.Visible = false;
                _toastTimer.Stop();
            };

            // Background update check on startup
            this.Shown += (s, e) =>
            {
                UpdateManager.CheckForUpdatesAsync(this, info =>
                {
                    if (info != null && info.HasUpdate)
                    {
                        if (_lblUpdateBadge != null)
                        {
                            _lblUpdateBadge.Text = "狩?v" + info.Version + " 媛??";
                            _lblUpdateBadge.Visible = true;
                        }
                        ShowToast("??踰꾩쟾 v" + info.Version + " ?낅뜲?댄듃媛 異쒖떆?섏뿀?듬땲??", ColorEmerald);
                    }
                }, silent: true);
            };

            // 1. Top Panel (Settings)
            Panel pnlSettings = new Panel
            {
                Dock = DockStyle.Top,
                Height = 110,
                BackColor = ColorBgCard,
                Padding = new Padding(15)
            };
            this.Controls.Add(pnlSettings);

            // Title inside Settings
            Label lblSettingsTitle = new Label
            {
                Text = "데이터베이스 연결 설정",
                Location = new Point(15, 10),
                Size = new Size(150, 20),
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold),
                ForeColor = ColorIndigo
            };
            pnlSettings.Controls.Add(lblSettingsTitle);

            // Demo Mode Checkbox
            _chkDemoMode = new CheckBox
            {
                Text = "가상 데이터 데모 모드",
                Location = new Point(175, 8),
                Size = new Size(160, 22),
                Checked = true,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                ForeColor = ColorTextMain
            };
            _chkDemoMode.CheckedChanged += ChkDemoMode_CheckedChanged;
            pnlSettings.Controls.Add(_chkDemoMode);

            // Status Badge
            _lblStatusBadge = new Label
            {
                Text = "데모 모드 (오프라인)",
                Location = new Point(345, 8),
                Size = new Size(130, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(245, 158, 11),
                ForeColor = Color.Black,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            pnlSettings.Controls.Add(_lblStatusBadge);

            // Maker Label
            Label lblMaker = new Label
            {
                Text = "만든이: 한솔인텍",
                Location = new Point(500, 8),
                Size = new Size(150, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlSettings.Controls.Add(lblMaker);

            _btnMyBox = new Button
            {
                Text = "MyBox",
                Location = new Point(0, 0),
                Size = new Size(76, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _btnMyBox.FlatAppearance.BorderSize = 0;
            _btnMyBox.Click += BtnMyBox_Click;
            pnlSettings.Controls.Add(_btnMyBox);

            _lblVersionBadge = new Label
            {
                Text = "v" + UpdateManager.CurrentVersion,
                Location = new Point(0, 0),
                Size = new Size(72, 24),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(30, 41, 59), // Slate 800
                ForeColor = Color.FromArgb(147, 197, 253), // Blue 300
                Font = new Font("留묒? 怨좊뵓", 9.0F, FontStyle.Bold)
            };
            pnlSettings.Controls.Add(_lblVersionBadge);

            _btnCheckUpdate = new Button
            {
                Text = "\uC5C5\uB370\uC774\uD2B8 \uD655\uC778",
                Location = new Point(0, 0),
                Size = new Size(125, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(79, 70, 229), // Indigo 600
                ForeColor = Color.White,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 8.5F, FontStyle.Bold)
            };
            _btnCheckUpdate.FlatAppearance.BorderSize = 0;
            _btnCheckUpdate.Click += (s, e) => UpdateManager.CheckForUpdatesAsync(this, null, silent: false);
            pnlSettings.Controls.Add(_btnCheckUpdate);

            _lblUpdateBadge = new Label
            {
                Text = "\u2B50 \uC0C8 \uBC84\uC804 \uAC00\uB2A5!",
                Location = new Point(0, 0),
                Size = new Size(115, 26),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(220, 38, 38), // Red 600
                ForeColor = Color.White,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 8.5F, FontStyle.Bold),
                Visible = false,
                Cursor = Cursors.Hand
            };
            _lblUpdateBadge.Click += (s, e) => UpdateManager.CheckForUpdatesAsync(this, null, silent: false);
            pnlSettings.Controls.Add(_lblUpdateBadge);

            // SQL Service Control
            _lblSqlServiceStatus = new Label
            {
                Text = "● SQL 서비스: 확인중...",
                Location = new Point(940, 8),
                Size = new Size(140, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            pnlSettings.Controls.Add(_lblSqlServiceStatus);

            _btnSqlServiceStart = new Button
            {
                Text = "▶ 시작",
                Location = new Point(1090, 6),
                Size = new Size(70, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _btnSqlServiceStart.FlatAppearance.BorderSize = 0;
            _btnSqlServiceStart.Click += (s, e) => ControlSqlService(true);
            pnlSettings.Controls.Add(_btnSqlServiceStart);

            _btnSqlServiceStop = new Button
            {
                Text = "■ 중지",
                Location = new Point(1165, 6),
                Size = new Size(70, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _btnSqlServiceStop.FlatAppearance.BorderSize = 0;
            _btnSqlServiceStop.Click += (s, e) => ControlSqlService(false);
            pnlSettings.Controls.Add(_btnSqlServiceStop);

            // Connection Settings Inner Panel
            Panel pnlSettingsFields = new Panel
            {
                Location = new Point(15, 38),
                Size = new Size(1240, 65)
            };
            pnlSettings.Controls.Add(pnlSettingsFields);

            // Server Label & Textbox
            Label lblServer = new Label { Text = "서버 주소", Location = new Point(0, 5), Size = new Size(60, 20), ForeColor = ColorTextSec };
            _txtServer = new TextBox { Location = new Point(0, 28), Size = new Size(150, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlSettingsFields.Controls.Add(lblServer);
            pnlSettingsFields.Controls.Add(_txtServer);

            // Windows Auth Checkbox
            _chkIntegratedSecurity = new CheckBox
            {
                Text = "윈도우 계정 연결 (Windows Auth)",
                Location = new Point(165, 28),
                Size = new Size(220, 25),
                ForeColor = ColorTextMain
            };
            _chkIntegratedSecurity.CheckedChanged += ChkIntegratedSecurity_CheckedChanged;
            pnlSettingsFields.Controls.Add(_chkIntegratedSecurity);

            // Username/Password Panel
            _pnlCredentials = new Panel { Location = new Point(390, 0), Size = new Size(290, 60) };
            pnlSettingsFields.Controls.Add(_pnlCredentials);

            Label lblUser = new Label { Text = "사용자 ID", Location = new Point(0, 5), Size = new Size(70, 20), ForeColor = ColorTextSec };
            _txtUser = new TextBox { Location = new Point(0, 28), Size = new Size(100, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _pnlCredentials.Controls.Add(lblUser);
            _pnlCredentials.Controls.Add(_txtUser);

            Label lblPassword = new Label { Text = "비밀번호", Location = new Point(115, 5), Size = new Size(70, 20), ForeColor = ColorTextSec };
            _txtPassword = new TextBox { Location = new Point(115, 28), Size = new Size(120, 25), PasswordChar = '●', BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _pnlCredentials.Controls.Add(lblPassword);
            _pnlCredentials.Controls.Add(_txtPassword);

            // DB Load Button
            _btnLoadDbs = new Button
            {
                Text = "DB 불러오기",
                Location = new Point(695, 25),
                Size = new Size(95, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White
            };
            _btnLoadDbs.FlatAppearance.BorderSize = 0;
            _btnLoadDbs.Click += BtnLoadDbs_Click;
            pnlSettingsFields.Controls.Add(_btnLoadDbs);

            // DB ComboBox
            Label lblDb = new Label { Text = "데이터베이스", Location = new Point(805, 5), Size = new Size(90, 20), ForeColor = ColorTextSec };
            _cmbDatabases = new ComboBox { Location = new Point(805, 28), Size = new Size(180, 25), DropDownStyle = ComboBoxStyle.DropDownList, BackColor = ColorBgMain, ForeColor = ColorTextMain };
            pnlSettingsFields.Controls.Add(lblDb);
            pnlSettingsFields.Controls.Add(_cmbDatabases);

            // Save Configuration Button (saves connection settings + layout positions)
            _btnSaveConfig = new Button
            {
                Text = "💾 설정/위치 저장",
                Location = new Point(990, 23),
                Size = new Size(125, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(99, 102, 241),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnSaveConfig.FlatAppearance.BorderSize = 0;
            _btnSaveConfig.Click += BtnSaveConfig_Click;
            pnlSettingsFields.Controls.Add(_btnSaveConfig);

            NormalizeSettingsHeader(
                pnlSettings,
                pnlSettingsFields,
                lblSettingsTitle,
                lblMaker,
                lblServer,
                lblDb);

            // 2. TabControl
            Panel mainTabHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain
            };
            this.Controls.Add(mainTabHost);
            mainTabHost.BringToFront();

            _tabControl = new TablessTabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            mainTabHost.Controls.Add(_tabControl);
            _tabControl.BringToFront();
            _tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

            // Tab 1: DB 응급 복구 및 무결성 검사 (맨 앞으로 배치)
            _tabDbRecovery = new TabPage
            {
                Text = "🚨 DB 응급 복구 및 무결성 검사",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabDbRecovery);

            // Tab 2: 조제고객 관련 관리
            _tabDispenseCustomerManagement = new TabPage
            {
                Text = "🏥 조제고객 관련 관리",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabDispenseCustomerManagement);

            Panel dispenseTabHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain
            };
            _tabDispenseCustomerManagement.Controls.Add(dispenseTabHost);

            _subTabDispenseCustomer = new TablessTabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            dispenseTabHost.Controls.Add(_subTabDispenseCustomer);

            // 서브 Tab 1: 차트번호 검증 및 조회
            _tabMainWorkspace = new TabPage
            {
                Text = "🔍 차트번호 검증 및 조회",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabMainWorkspace);

            // 서브 Tab 2: 차트 오류 해결 도구
            _tabTroubleshooter = new TabPage
            {
                Text = "🛠️ 차트 오류 해결 도구",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabTroubleshooter);

            // 서브 Tab 3: 백업로그청구기반검사 (차트오류해결도구 바로 옆)
            _tabLogClaimMismatch = new TabPage
            {
                Text = "🔬 백업로그청구기반검사",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabLogClaimMismatch);

            // 서브 Tab 4: 과거 이력 관리
            _tabPastHistoryManagement = new TabPage
            {
                Text = "📂 과거 이력 관리",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabPastHistoryManagement);


            // 서브 Tab 4: 청구 누락 점검
            _tabClaimComparison = new TabPage
            {
                Text = "📊 청구 누락 점검",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabClaimComparison);

            // 서브 Tab 5: 처방전 내역 삭제 (탭 생성만 하고 초기화는 아래 InitializePrescriptionDeleteTab에서)
            _tabPrescriptionDelete = new TabPage
            {
                Text = "❌ 처방전 내역 삭제",
                BackColor = ColorBgMain
            };
            _subTabDispenseCustomer.TabPages.Add(_tabPrescriptionDelete);
            
            // Tab 3: 의사면허 중복 관리
            _tabDoctorLicense = new TabPage
            {
                Text = "👨‍⚕️ 의사면허 중복 관리",
                BackColor = ColorBgMain
            };
            // _tabControl.TabPages.Add(_tabDoctorLicense); // 기초 데이터 관리 서브 탭에 추가하므로 여기선 제외

            Panel pnlDocTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };

            Label lblDocTitle = new Label
            {
                Text = "요양기관 내에 등록된 의사면허번호(DC_ID)가 중복 생성된 내역을 조회하고 일괄 정리합니다.",
                Location = new Point(15, 20),
                Size = new Size(580, 25),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Italic)
            };
            pnlDocTop.Controls.Add(lblDocTitle);

            _btnSearchDoctors = new Button
            {
                Text = "🔍 중복 의사면허 조회",
                Location = new Point(610, 15),
                Size = new Size(160, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnSearchDoctors.FlatAppearance.BorderSize = 0;
            _btnSearchDoctors.Click += BtnSearchDoctors_Click;
            pnlDocTop.Controls.Add(_btnSearchDoctors);

            _btnDeleteDoctors = new Button
            {
                Text = "🗑️ 중복 의사면허 일괄 삭제",
                Location = new Point(780, 15),
                Size = new Size(190, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnDeleteDoctors.FlatAppearance.BorderSize = 0;
            _btnDeleteDoctors.Click += BtnDeleteDoctors_Click;
            pnlDocTop.Controls.Add(_btnDeleteDoctors);

            _lblDoctorStatus = new Label
            {
                Text = "",
                Location = new Point(980, 20),
                Size = new Size(150, 25),
                ForeColor = ColorEmerald,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlDocTop.Controls.Add(_lblDoctorStatus);
            NormalizeDoctorHeader(pnlDocTop, lblDocTitle);

            _tabDoctorLicense.Controls.Add(pnlDocTop);

            _dgvDoctors = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvDoctors.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvDoctors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvDoctors.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvDoctors.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvDoctors.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvDoctors.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvDoctors.DefaultCellStyle.SelectionForeColor = Color.White;

            _tabDoctorLicense.Controls.Add(_dgvDoctors);
            _dgvDoctors.BringToFront();

            // DB 응급 복구 및 무결성 검사 탭은 맨 앞에서 생성되어 추가되었습니다.

            Panel pnlRecTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 110,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };

            Label lblRecTitle = new Label
            {
                Text = "⚠️ 데이터베이스가 응급 모드(EMERGENCY)에 빠졌을 때 일괄적으로 무결성 검사 및 데이터 복구 작업을 실행합니다.",
                Location = new Point(15, 12),
                Size = new Size(800, 20),
                ForeColor = ColorAlarm,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlRecTop.Controls.Add(lblRecTitle);

            Label lblRecSub = new Label
            {
                Text = "※ 주의: REPAIR_ALLOW_DATA_LOSS 복구 과정에서 일부 손상된 데이터가 제거(유실)될 가능성이 있으므로 신중히 진행해야 합니다.",
                Location = new Point(15, 32),
                Size = new Size(800, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Italic)
            };
            pnlRecTop.Controls.Add(lblRecSub);

            // DB Checkboxes
            _chkDbPmMain = new CheckBox { Text = "pm_main", Location = new Point(20, 65), Size = new Size(110, 25), Checked = true, ForeColor = ColorTextMain };
            _chkDbPmplusDums = new CheckBox { Text = "pmplus_dums", Location = new Point(130, 65), Size = new Size(130, 25), Checked = true, ForeColor = ColorTextMain };
            _chkDbPmplusImage = new CheckBox { Text = "pmplus_image", Location = new Point(265, 65), Size = new Size(130, 25), Checked = true, ForeColor = ColorTextMain };
            _chkDbPmplusJoblog = new CheckBox { Text = "pmplus_joblog", Location = new Point(400, 65), Size = new Size(140, 25), Checked = true, ForeColor = ColorTextMain };

            pnlRecTop.Controls.Add(_chkDbPmMain);
            pnlRecTop.Controls.Add(_chkDbPmplusDums);
            pnlRecTop.Controls.Add(_chkDbPmplusImage);
            pnlRecTop.Controls.Add(_chkDbPmplusJoblog);

            _btnRunDbRecovery = new Button
            {
                Text = "⚡ DB 응급 복구 실행",
                Location = new Point(570, 60),
                Size = new Size(180, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnRunDbRecovery.FlatAppearance.BorderSize = 0;
            _btnRunDbRecovery.Click += BtnRunDbRecovery_Click;
            pnlRecTop.Controls.Add(_btnRunDbRecovery);

            _btnShrinkDb = new Button
            {
                Text = "📦 DB 축소",
                Location = new Point(760, 60),
                Size = new Size(120, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(79, 70, 229), // Indigo 600
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnShrinkDb.FlatAppearance.BorderSize = 0;
            _btnShrinkDb.Click += BtnShrinkDb_Click;
            pnlRecTop.Controls.Add(_btnShrinkDb);

            _btnShrinkLog = new Button
            {
                Text = "📄 LOG 축소",
                Location = new Point(890, 60),
                Size = new Size(120, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(13, 148, 136), // Teal 600
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnShrinkLog.FlatAppearance.BorderSize = 0;
            _btnShrinkLog.Click += BtnShrinkLog_Click;
            pnlRecTop.Controls.Add(_btnShrinkLog);

            _btnDropDrugUpdateDb = new Button
            {
                Text = "PM_DRUGUPDATE 삭제",
                Location = new Point(740, 18),
                Size = new Size(155, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.0F, FontStyle.Bold)
            };
            _btnDropDrugUpdateDb.FlatAppearance.BorderSize = 0;
            _btnDropDrugUpdateDb.Click += BtnDropDrugUpdateDb_Click;
            pnlRecTop.Controls.Add(_btnDropDrugUpdateDb);

            _btnDropDurUpdateDb = new Button
            {
                Text = "PM_DURUPDATE 삭제",
                Location = new Point(905, 18),
                Size = new Size(155, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.0F, FontStyle.Bold)
            };
            _btnDropDurUpdateDb.FlatAppearance.BorderSize = 0;
            _btnDropDurUpdateDb.Click += BtnDropDurUpdateDb_Click;
            pnlRecTop.Controls.Add(_btnDropDurUpdateDb);

            NormalizeDbRecoveryHeader(pnlRecTop);

            _tabDbRecovery.Controls.Add(pnlRecTop);

            // Log Console (TextBox)
            _txtDbRecoveryLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(15, 23, 42), // Slate 900
                ForeColor = Color.FromArgb(34, 197, 94), // Green 500
                Font = new Font("Consolas", 9.5F),
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            _tabDbRecovery.Controls.Add(_txtDbRecoveryLog);
            _txtDbRecoveryLog.BringToFront();

            // 2. Main Workspace SplitContainer
            SplitContainer splitMain = new SplitContainer();
            splitMain.Dock = DockStyle.Fill;
            splitMain.Orientation = Orientation.Vertical;
            splitMain.Size = new Size(1100, 600);
            splitMain.Panel1MinSize = 450;
            splitMain.Panel2MinSize = 450;
            splitMain.SplitterDistance = 600;
            splitMain.BackColor = ColorBorder;
            _tabMainWorkspace.Controls.Add(splitMain);

            // 2-1. Left Panel (Search & Grid results)
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = ColorBgMain, Padding = new Padding(15) };
            splitMain.Panel1.Controls.Add(pnlLeft);

            // Step 1 Header
            Panel pnlLeftHeader = new Panel { Dock = DockStyle.Top, Height = 40 };
            Label lblStep1 = new Label { Text = "Step 1", Location = new Point(0, 5), Size = new Size(50, 22), BackColor = Color.FromArgb(30, 58, 138), ForeColor = Color.FromArgb(191, 219, 254), TextAlign = ContentAlignment.MiddleCenter, Font = new Font("맑은 고딕", 8F, FontStyle.Bold) };
            Label lblLeftTitle = new Label { Text = "처방 리스트에서 차트번호 조회", Location = new Point(60, 5), Size = new Size(400, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold) };
            pnlLeftHeader.Controls.Add(lblStep1);
            pnlLeftHeader.Controls.Add(lblLeftTitle);

            // Search Inputs Panel
            Panel pnlSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };

            Label lblSearchName = new Label { Text = "환자 이름", Location = new Point(12, 10), Size = new Size(70, 20), ForeColor = ColorTextSec };
            _txtSearchName = new TextBox { Location = new Point(12, 33), Size = new Size(130, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlSearch.Controls.Add(lblSearchName);
            pnlSearch.Controls.Add(_txtSearchName);

            Label lblSearchJumin = new Label { Text = "주민번호 앞 7자리", Location = new Point(160, 10), Size = new Size(120, 20), ForeColor = ColorTextSec };
            _txtSearchJumin = new TextBox { Location = new Point(160, 33), Size = new Size(150, 25), MaxLength = 8, BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _txtSearchJumin.TextChanged += TxtSearchJumin_TextChanged;
            pnlSearch.Controls.Add(lblSearchJumin);
            pnlSearch.Controls.Add(_txtSearchJumin);

            _btnSearchRx = new Button
            {
                Text = "처방 검색",
                Location = new Point(330, 28),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnSearchRx.FlatAppearance.BorderSize = 0;
            _btnSearchRx.Click += BtnSearchRx_Click;
            pnlSearch.Controls.Add(_btnSearchRx);

            // Spacer Panel
            Panel pnlSpacer1 = new Panel { Dock = DockStyle.Top, Height = 15 };

            // DataGridView Results
            _dgvResults = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32
            };
            _dgvResults.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvResults.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvResults.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvResults.ColumnHeadersDefaultCellStyle.Padding = new Padding(6, 4, 6, 4);
            _dgvResults.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvResults.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvResults.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvResults.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvResults.CellClick += DgvResults_CellClick;
            _dgvResults.CellFormatting += DgvResults_CellFormatting;

            pnlLeft.Controls.Add(_dgvResults);
            pnlLeft.Controls.Add(pnlSpacer1);
            pnlLeft.Controls.Add(pnlSearch);
            pnlLeft.Controls.Add(pnlLeftHeader);

            // 2-2. Right Panel (Customer Details & verification)
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, BackColor = ColorBgMain, Padding = new Padding(15) };
            splitMain.Panel2.Controls.Add(pnlRight);

            // Step 2 Header
            Panel pnlRightHeader = new Panel { Dock = DockStyle.Top, Height = 40 };
            Label lblStep2 = new Label { Text = "Step 2", Location = new Point(0, 5), Size = new Size(50, 22), BackColor = Color.FromArgb(6, 95, 70), ForeColor = Color.FromArgb(167, 243, 208), TextAlign = ContentAlignment.MiddleCenter, Font = new Font("맑은 고딕", 8F, FontStyle.Bold) };
            Label lblRightTitle = new Label { Text = "차트번호 고객 정보 검증 (조회 및 대조)", Location = new Point(60, 5), Size = new Size(400, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold) };
            pnlRightHeader.Controls.Add(lblStep2);
            pnlRightHeader.Controls.Add(lblRightTitle);

            // Search ChartNo Panel
            Panel pnlSearchCust = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };

            Label lblSearchChrtNo = new Label { Text = "차트번호 직접 입력", Location = new Point(12, 10), Size = new Size(130, 20), ForeColor = ColorTextSec };
            _txtSearchChrtNo = new TextBox { Location = new Point(12, 33), Size = new Size(130, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlSearchCust.Controls.Add(lblSearchChrtNo);
            pnlSearchCust.Controls.Add(_txtSearchChrtNo);

            _btnSearchCust = new Button
            {
                Text = "고객 조회",
                Location = new Point(155, 28),
                Size = new Size(85, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnSearchCust.FlatAppearance.BorderSize = 0;
            _btnSearchCust.Click += BtnSearchCust_Click;
            pnlSearchCust.Controls.Add(_btnSearchCust);

            _btnRestoreCust = new Button
            {
                Text = "고객 복구",
                Location = new Point(250, 28),
                Size = new Size(85, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(245, 158, 11), // Amber 500
                ForeColor = Color.Black,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnRestoreCust.FlatAppearance.BorderSize = 0;
            _btnRestoreCust.Click += BtnRestoreCust_Click;
            pnlSearchCust.Controls.Add(_btnRestoreCust);

            // Toast Alert Label in Step 2 panel
            _lblToast = new Label
            {
                Location = new Point(345, 28),
                Size = new Size(90, 32),
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = ColorEmerald,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };
            pnlSearchCust.Controls.Add(_lblToast);

            // Spacer Panel
            Panel pnlSpacer2 = new Panel { Dock = DockStyle.Top, Height = 15 };

            // Detail Card Panel
            Panel pnlDetailCard = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(15)
            };

            // Customer Name & Chart Badge layout
            Panel pnlCustTitle = new Panel { Dock = DockStyle.Top, Height = 35 };
            _lblCustNameTitle = new Label { Text = "선택된 고객 없음", Location = new Point(0, 0), Size = new Size(150, 30), Font = new Font("맑은 고딕", 12F, FontStyle.Bold), ForeColor = Color.White };
            _lblCustChrtNoBadge = new Label { Text = "-", Location = new Point(160, 4), Size = new Size(120, 22), TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.FromArgb(6, 95, 70), ForeColor = Color.FromArgb(167, 243, 208), Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            pnlCustTitle.Controls.Add(_lblCustNameTitle);
            pnlCustTitle.Controls.Add(_lblCustChrtNoBadge);

            // Grid Layout for details
            Panel pnlCustGrid = new Panel { Dock = DockStyle.Top, Height = 150, Padding = new Padding(0, 10, 0, 0) };

            int ly = 15;
            Label l1 = new Label { Text = "차트번호:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustChrtNo = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(200, 20), ForeColor = ColorTextMain };
            pnlCustGrid.Controls.Add(l1); pnlCustGrid.Controls.Add(_lblCustChrtNo);

            ly += 25;
            Label l2 = new Label { Text = "등록 이름:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustName = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(200, 20), ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            pnlCustGrid.Controls.Add(l2); pnlCustGrid.Controls.Add(_lblCustName);

            ly += 25;
            Label l3 = new Label { Text = "주민번호:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustJumin = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(200, 20), ForeColor = ColorTextMain };
            pnlCustGrid.Controls.Add(l3); pnlCustGrid.Controls.Add(_lblCustJumin);

            ly += 25;
            Label l4 = new Label { Text = "연락처:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustPhone = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(200, 20), ForeColor = ColorTextMain };
            pnlCustGrid.Controls.Add(l4); pnlCustGrid.Controls.Add(_lblCustPhone);

            ly += 25;
            Label l5 = new Label { Text = "주소:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustAddress = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(400, 20), ForeColor = ColorTextMain };
            pnlCustGrid.Controls.Add(l5); pnlCustGrid.Controls.Add(_lblCustAddress);

            ly += 25;
            Label l6 = new Label { Text = "최초방문:", Location = new Point(0, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _lblCustFirstVisit = new Label { Text = "-", Location = new Point(90, ly), Size = new Size(200, 20), ForeColor = ColorTextMain };
            pnlCustGrid.Controls.Add(l6); pnlCustGrid.Controls.Add(_lblCustFirstVisit);

            // History Label
            Label lblHistHeader = new Label { Text = "최근 처방 기록 (최대 10건)", Dock = DockStyle.Top, Height = 25, ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };

            // History ListBox
            _lstRxHistory = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                BorderStyle = BorderStyle.FixedSingle
            };

            // Add inner controls in reverse Z-order to layout properly without overlaps
            pnlDetailCard.Controls.Add(_lstRxHistory);
            pnlDetailCard.Controls.Add(lblHistHeader);
            pnlDetailCard.Controls.Add(pnlCustGrid);
            pnlDetailCard.Controls.Add(pnlCustTitle);

            // Add right panel controls in reverse Z-order so they dock top-down correctly
            pnlRight.Controls.Add(pnlDetailCard);
            pnlRight.Controls.Add(pnlSpacer2);
            pnlRight.Controls.Add(pnlSearchCust);
            pnlRight.Controls.Add(pnlRightHeader);

            // Initialize 5th Tab (Data Management)
            InitializeDataManagementTab();

            // Initialize monthly dispensing/claim comparison
            InitializeClaimComparisonTab();

            // Initialize Prescription Delete
            InitializePrescriptionDeleteTab();

            // Initialize Log & Claim Mismatch Scanner Tab
            InitializeLogClaimMismatchTab();

            // Initialize Past History Management Tab
            InitializePastHistoryTab();

            // Initialize Narcotics Management Tab
            InitializeNarcoticsManagementTab();

            // Initialize Query Runner Tab
            InitializeQueryRunnerTab();

            StripParenthesizedTabNames(_subTabDataManagement);
            InstallVisibleTabNavigation(mainTabHost, _tabControl);
            InstallVisibleTabNavigation(dispenseTabHost, _subTabDispenseCustomer);
        }

        private void NormalizeSettingsHeader(
            Panel pnlSettings,
            Panel pnlSettingsFields,
            Label lblSettingsTitle,
            Label lblMaker,
            Label lblServer,
            Label lblDb)
        {
            pnlSettings.Height = 136;
            pnlSettings.Padding = new Padding(16, 8, 16, 10);

            lblSettingsTitle.Location = new Point(16, 12);
            lblSettingsTitle.Size = new Size(150, 22);

            _chkDemoMode.Location = new Point(176, 10);
            _chkDemoMode.Size = new Size(165, 24);

            _lblStatusBadge.Location = new Point(348, 9);
            _lblStatusBadge.Size = new Size(148, 26);

            lblMaker.Location = new Point(510, 10);
            lblMaker.Size = new Size(108, 24);

            if (_btnMyBox != null)
            {
                _btnMyBox.Location = new Point(lblMaker.Right + 8, 9);
                _btnMyBox.Size = new Size(76, 26);
                _btnMyBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _btnMyBox.BringToFront();
            }

            if (_lblVersionBadge != null)
            {
                _lblVersionBadge.Location = new Point(lblMaker.Right + 8, 10);
                _lblVersionBadge.Size = new Size(68, 24);
                _lblVersionBadge.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _lblVersionBadge.BringToFront();
            }

            if (_btnMyBox != null)
            {
                _btnMyBox.Location = new Point(_lblVersionBadge != null ? _lblVersionBadge.Right + 8 : lblMaker.Right + 8, 9);
                _btnMyBox.Size = new Size(76, 26);
                _btnMyBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _btnMyBox.BringToFront();
            }

            if (_btnCheckUpdate != null)
            {
                _btnCheckUpdate.Location = new Point(_btnMyBox != null ? _btnMyBox.Right + 8 : lblMaker.Right + 80, 9);
                _btnCheckUpdate.Size = new Size(GetButtonContentWidth(_btnCheckUpdate), 26);
                _btnCheckUpdate.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _btnCheckUpdate.BringToFront();
            }

            if (_lblUpdateBadge != null)
            {
                _lblUpdateBadge.Location = new Point(_btnCheckUpdate != null ? _btnCheckUpdate.Right + 8 : lblMaker.Right + 120, 9);
                _lblUpdateBadge.Size = new Size(115, 26);
                _lblUpdateBadge.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _lblUpdateBadge.BringToFront();
            }

            pnlSettingsFields.Location = new Point(16, 48);
            pnlSettingsFields.Height = 76;
            pnlSettingsFields.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            pnlSettingsFields.Width = pnlSettings.ClientSize.Width - 32;

            Action layout = delegate
            {
                LayoutSettingsHeaderControls(pnlSettings, pnlSettingsFields, lblServer, lblDb);
            };

            pnlSettings.Resize += delegate { layout(); };
            layout();
        }

        private void LayoutSettingsHeaderControls(
            Panel pnlSettings,
            Panel pnlSettingsFields,
            Label lblServer,
            Label lblDb)
        {
            int panelWidth = Math.Max(0, pnlSettings.ClientSize.Width);
            pnlSettingsFields.Width = Math.Max(0, panelWidth - 32);

            int serviceButtonTop = 8;
            int serviceButtonWidth = Math.Max(GetButtonContentWidth(_btnSqlServiceStart), GetButtonContentWidth(_btnSqlServiceStop));
            int serviceButtonGap = 8;
            int serviceRightMargin = 16;
            _btnSqlServiceStop.Size = new Size(serviceButtonWidth, 30);
            _btnSqlServiceStop.Location = new Point(panelWidth - serviceRightMargin - _btnSqlServiceStop.Width, serviceButtonTop);
            _btnSqlServiceStop.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            _btnSqlServiceStart.Size = new Size(serviceButtonWidth, 30);
            _btnSqlServiceStart.Location = new Point(_btnSqlServiceStop.Left - serviceButtonGap - _btnSqlServiceStart.Width, serviceButtonTop);
            _btnSqlServiceStart.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            _lblSqlServiceStatus.Size = new Size(170, 24);
            _lblSqlServiceStatus.Location = new Point(_btnSqlServiceStart.Left - 182, serviceButtonTop + 3);
            _lblSqlServiceStatus.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            int w = Math.Max(0, pnlSettingsFields.ClientSize.Width);
            int loadButtonWidth = GetButtonContentWidth(_btnLoadDbs);
            int saveButtonWidth = GetButtonContentWidth(_btnSaveConfig);
            int dbRowRequiredWidth = loadButtonWidth + 18 + 170 + 18 + saveButtonWidth;
            bool compact = w < 1160;
            bool dbWrap = w < (700 + dbRowRequiredWidth);
            int y1 = 2;
            int yInput1 = 27;
            int y2 = compact ? 70 : 2;
            int yInput2 = compact ? 95 : 27;
            int y3 = 132;
            int yInput3 = 157;

            pnlSettings.Height = dbWrap ? 240 : (compact ? 196 : 136);
            pnlSettingsFields.Height = dbWrap ? 182 : (compact ? 138 : 76);

            lblServer.Location = new Point(0, y1);
            lblServer.Size = new Size(80, 20);
            _txtServer.Location = new Point(0, yInput1);
            _txtServer.Size = new Size(150, 26);
            _txtServer.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            _chkIntegratedSecurity.Location = new Point(166, yInput1);
            _chkIntegratedSecurity.Size = new Size(220, 26);
            _chkIntegratedSecurity.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            _pnlCredentials.Location = new Point(404, 0);
            _pnlCredentials.Size = new Size(270, 62);
            _pnlCredentials.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            _txtUser.Size = new Size(104, 26);
            _txtPassword.Location = new Point(118, 28);
            _txtPassword.Size = new Size(126, 26);

            if (dbWrap)
            {
                _btnLoadDbs.Location = new Point(0, yInput3 - 2);
                _btnLoadDbs.Size = new Size(loadButtonWidth, 32);

                int dbX = _btnLoadDbs.Right + 18;
                lblDb.Location = new Point(dbX, y3);
                lblDb.Size = new Size(105, 20);
                _cmbDatabases.Location = new Point(dbX, yInput3);
                _cmbDatabases.Size = new Size(Math.Max(170, w - dbX - saveButtonWidth - 36), 26);

                _btnSaveConfig.Location = new Point(w - saveButtonWidth, yInput3 - 2);
                _btnSaveConfig.Size = new Size(saveButtonWidth, 32);
            }
            else if (compact)
            {
                _btnLoadDbs.Location = new Point(0, yInput2 - 2);
                _btnLoadDbs.Size = new Size(loadButtonWidth, 32);

                int compactDbX = _btnLoadDbs.Right + 18;
                lblDb.Location = new Point(compactDbX, y2);
                lblDb.Size = new Size(105, 20);
                _cmbDatabases.Location = new Point(compactDbX, yInput2);
                _cmbDatabases.Size = new Size(Math.Max(170, w - compactDbX - saveButtonWidth - 36), 26);

                _btnSaveConfig.Location = new Point(w - saveButtonWidth, yInput2 - 2);
                _btnSaveConfig.Size = new Size(saveButtonWidth, 32);
            }
            else
            {
                int saveX = Math.Max(0, w - saveButtonWidth);
                int comboX = Math.Max(0, saveX - 188);
                int loadX = Math.Max(0, comboX - loadButtonWidth - 20);

                _btnLoadDbs.Location = new Point(loadX, yInput1 - 2);
                _btnLoadDbs.Size = new Size(loadButtonWidth, 32);

                lblDb.Location = new Point(comboX, y1);
                lblDb.Size = new Size(105, 20);
                _cmbDatabases.Location = new Point(comboX, yInput1);
                _cmbDatabases.Size = new Size(170, 26);

                _btnSaveConfig.Location = new Point(saveX, yInput1 - 2);
                _btnSaveConfig.Size = new Size(saveButtonWidth, 32);
            }

            _btnLoadDbs.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            lblDb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _cmbDatabases.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _btnSaveConfig.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        }

        private void BtnMyBox_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://naver.me/xUwhAChe");
            }
            catch (Exception ex)
            {
                MessageBox.Show("MyBox 웹페이지를 열지 못했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void NormalizeDbRecoveryHeader(Panel pnlRecTop)
        {
            Action layout = delegate
            {
                int yTop = 18;
                int yBottom = 60;
                int gap = 10;

                if (_btnDropDrugUpdateDb != null) _btnDropDrugUpdateDb.Size = new Size(GetButtonContentWidth(_btnDropDrugUpdateDb), 32);
                if (_btnDropDurUpdateDb != null) _btnDropDurUpdateDb.Size = new Size(GetButtonContentWidth(_btnDropDurUpdateDb), 32);

                _btnRunDbRecovery.Size = new Size(GetButtonContentWidth(_btnRunDbRecovery), 32);
                _btnShrinkDb.Size = new Size(GetButtonContentWidth(_btnShrinkDb), 32);
                _btnShrinkLog.Size = new Size(GetButtonContentWidth(_btnShrinkLog), 32);

                int bottomButtonsWidth = _btnRunDbRecovery.Width + _btnShrinkDb.Width + _btnShrinkLog.Width + (gap * 2);
                int topButtonsWidth = (_btnDropDrugUpdateDb != null ? _btnDropDrugUpdateDb.Width : 140) + (_btnDropDurUpdateDb != null ? _btnDropDurUpdateDb.Width : 140) + gap;

                int maxBlockWidth = Math.Max(bottomButtonsWidth, topButtonsWidth);
                int startX = Math.Max(520, pnlRecTop.ClientSize.Width - maxBlockWidth - 30);

                // Bottom row: DB ?묎툒 蹂듦뎄 ?ㅽ뻾, DB 異뺤냼, LOG 異뺤냼
                _btnRunDbRecovery.Location = new Point(startX, yBottom);
                _btnShrinkDb.Location = new Point(_btnRunDbRecovery.Right + gap, yBottom);
                _btnShrinkLog.Location = new Point(_btnShrinkDb.Right + gap, yBottom);

                // Top row: PM_DRUGUPDATE ??젣, PM_DURUPDATE ??젣
                // ?ㅻⅨ履???_btnShrinkLog.Right)??湲곗??쇰줈 ?쇱そ?쇰줈 ??같移섑븯???붾㈃ 諛뽰쑝濡??덈? ?섍?吏 ?딅룄濡??꾨꼍 蹂댁옣!
                if (_btnDropDrugUpdateDb != null && _btnDropDurUpdateDb != null)
                {
                    int rightLimit = _btnShrinkLog.Right;
                    _btnDropDurUpdateDb.Location = new Point(rightLimit - _btnDropDurUpdateDb.Width, yTop);
                    _btnDropDrugUpdateDb.Location = new Point(_btnDropDurUpdateDb.Left - gap - _btnDropDrugUpdateDb.Width, yTop);

                    _btnDropDrugUpdateDb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                    _btnDropDurUpdateDb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                }

                _btnRunDbRecovery.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _btnShrinkDb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                _btnShrinkLog.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            };

            pnlRecTop.Resize += delegate { layout(); };
            layout();
        }
        private void NormalizeDoctorHeader(Panel pnlDocTop, Label lblDocTitle)
        {
            Action layout = delegate
            {
                int gap = 10;
                int rightPadding = 15;
                int buttonTop = 15;

                _btnSearchDoctors.Size = new Size(GetButtonContentWidth(_btnSearchDoctors), 30);
                _btnDeleteDoctors.Size = new Size(GetButtonContentWidth(_btnDeleteDoctors), 30);

                int statusWidth = Math.Max(120, TextRenderer.MeasureText(_lblDoctorStatus.Text ?? "", _lblDoctorStatus.Font).Width + 20);
                _lblDoctorStatus.Size = new Size(statusWidth, 30);

                int totalWidth = _btnSearchDoctors.Width + _btnDeleteDoctors.Width + _lblDoctorStatus.Width + (gap * 2);
                int startX = pnlDocTop.ClientSize.Width - totalWidth - rightPadding;
                if (startX < 15)
                {
                    pnlDocTop.Height = 96;
                    startX = 15;
                    buttonTop = 55;
                    lblDocTitle.Width = Math.Max(200, pnlDocTop.ClientSize.Width - 30);
                }
                else
                {
                    pnlDocTop.Height = 60;
                    lblDocTitle.Width = Math.Max(200, startX - 30);
                }

                _btnSearchDoctors.Location = new Point(startX, buttonTop);
                _btnDeleteDoctors.Location = new Point(_btnSearchDoctors.Right + gap, buttonTop);
                _lblDoctorStatus.Location = new Point(_btnDeleteDoctors.Right + gap, buttonTop);
                _lblDoctorStatus.TextAlign = ContentAlignment.MiddleLeft;
            };

            pnlDocTop.Resize += delegate { layout(); };
            _lblDoctorStatus.TextChanged += delegate { layout(); };
            layout();
        }

        private void NormalizeDataManagementSplit(SplitContainer split)
        {
            int ignored = split != null ? split.SplitterDistance : 0;
            NormalizeDataManagementSplit(split, ref ignored);
        }

        private void NormalizeDataManagementSplit(SplitContainer split, ref int storedDistance)
        {
            if (split == null || split.ClientSize.Width <= 0) return;

            int targetRightMin = 460;
            int leftMin = 280;
            int rightMin = split.ClientSize.Width >= leftMin + targetRightMin
                ? targetRightMin
                : Math.Max(120, split.ClientSize.Width - leftMin);
            int maxDistance = split.ClientSize.Width - rightMin;
            if (maxDistance < split.Panel1MinSize) maxDistance = split.Panel1MinSize;

            int desired = storedDistance > 0 ? storedDistance : split.SplitterDistance;
            int distance = Math.Min(desired, maxDistance);
            distance = Math.Max(split.Panel1MinSize, distance);

            try
            {
                split.SplitterDistance = distance;
                storedDistance = distance;
            }
            catch { }
        }

        private void NormalizeRightPanelSplit(SplitContainer split, ref int storedDistance, int rightMinWidth, int leftMinWidth)
        {
            if (split == null || split.ClientSize.Width <= 0) return;

            int width = split.ClientSize.Width;
            int available = width - split.SplitterWidth - 2;
            if (available < 240) return;

            int rightMin = Math.Min(rightMinWidth, Math.Max(120, available - leftMinWidth));
            int leftMin = Math.Min(leftMinWidth, Math.Max(120, available - rightMin));

            if (leftMin + rightMin > available)
            {
                leftMin = Math.Max(120, available - rightMin);
            }

            if (leftMin + rightMin > available)
            {
                rightMin = Math.Max(120, available - leftMin);
            }

            if (leftMin + rightMin > available) return;

            try
            {
                split.Panel1MinSize = leftMin;
                split.Panel2MinSize = rightMin;
            }
            catch { return; }

            int maxDistance = width - split.Panel2MinSize;
            int desired = storedDistance > 0 ? storedDistance : split.SplitterDistance;
            int distance = Math.Min(desired, maxDistance);
            distance = Math.Max(split.Panel1MinSize, distance);

            try
            {
                split.SplitterDistance = distance;
                storedDistance = distance;
            }
            catch { }
        }

        private void NormalizeRightPanelSplit(SplitContainer split, int rightMinWidth, int leftMinWidth)
        {
            int ignored = split != null ? split.SplitterDistance : 0;
            NormalizeRightPanelSplit(split, ref ignored, rightMinWidth, leftMinWidth);
        }

        private void ApplyContentSizedColumns(DataGridView dgv)
        {
            if (dgv == null || dgv.Columns.Count == 0) return;

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Font headerFont = dgv.ColumnHeadersDefaultCellStyle.Font ?? dgv.Font;
            Font cellFont = dgv.DefaultCellStyle.Font ?? dgv.Font;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Visible) continue;

                int targetWidth = TextRenderer.MeasureText(col.HeaderText ?? "", headerFont).Width + 32;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    object value = row.Cells[col.Index].FormattedValue;
                    string text = value == null ? "" : value.ToString();
                    int cellWidth = TextRenderer.MeasureText(text, cellFont).Width + 28;
                    if (cellWidth > targetWidth) targetWidth = cellWidth;
                }

                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                col.MinimumWidth = Math.Min(Math.Max(60, targetWidth), 180);
                col.Width = Math.Min(Math.Max(60, targetWidth), 360);
            }
        }

        private void StripParenthesizedTabNames(TabControl tabControl)
        {
            if (tabControl == null) return;

            foreach (TabPage page in tabControl.TabPages)
            {
                string text = page.Text ?? "";
                int open = text.IndexOf('(');
                int close = text.LastIndexOf(')');
                if (open >= 0 && close > open)
                {
                    page.Text = text.Substring(0, open).TrimEnd();
                }
            }
        }

        // --- Event Handlers & Business Logic ---

        private void ChkDemoMode_CheckedChanged(object sender, EventArgs e)
        {
            ToggleDemoMode(_chkDemoMode.Checked);
        }

        private void ToggleDemoMode(bool isDemo)
        {
            _isDemo = isDemo;
            if (isDemo)
            {
                _lblStatusBadge.Text = "데모 모드 (가상)";
                _lblStatusBadge.BackColor = Color.FromArgb(245, 158, 11);
                _lblStatusBadge.ForeColor = Color.Black;

                // Disable DB Credentials Input fields in Demo Mode
                _txtServer.Enabled = false;
                _chkIntegratedSecurity.Enabled = false;
                _txtUser.Enabled = false;
                _txtPassword.Enabled = false;
                _btnLoadDbs.Enabled = false;
                _cmbDatabases.Enabled = false;
                UpdateStatus("데모 모드 - 실제 DB를 변경하지 않습니다. 검색과 복구 흐름을 연습할 수 있습니다.");
            }
            else
            {
                _lblStatusBadge.Text = "실서버 모드 (연결대기)";
                _lblStatusBadge.BackColor = ColorBorder;
                _lblStatusBadge.ForeColor = ColorTextMain;

                // Enable DB settings
                _txtServer.Enabled = true;
                _chkIntegratedSecurity.Enabled = true;
                _btnLoadDbs.Enabled = true;
                _cmbDatabases.Enabled = true;
                ChkIntegratedSecurity_CheckedChanged(null, null);
                UpdateStatus("실서버 모드 - DB 목록을 불러온 뒤 조회/작업 대상을 반드시 확인하세요.");
            }

            if (_troubleshooter != null)
            {
                _troubleshooter.ToggleDemoMode(isDemo);
            }

            if (_cmbQueryDbSelector != null && _cmbQueryDbSelector.Items.Count > 0)
            {
                LoadQueryRunnerDatabases();
            }

            RefreshAttachedBackupStatus();
        }

        private void ChkIntegratedSecurity_CheckedChanged(object sender, EventArgs e)
        {
            if (_chkIntegratedSecurity.Checked)
            {
                _pnlCredentials.Enabled = false;
            }
            else
            {
                _pnlCredentials.Enabled = true;
            }
        }

        private void TxtSearchJumin_TextChanged(object sender, EventArgs e)
        {
            // Auto hyphen format: YYMMDD-G
            string text = _txtSearchJumin.Text.Replace("-", "");
            if (text.Length > 6)
            {
                _txtSearchJumin.Text = text.Substring(0, 6) + "-" + text.Substring(6, Math.Min(1, text.Length - 6));
                _txtSearchJumin.SelectionStart = _txtSearchJumin.Text.Length;
            }
        }

        // Load Configuration from config file next to exe (portable)
        private const string DefaultServerAddress = @".\pmplus20";
        private const string DefaultDatabaseName = "PM_MAIN";

        private string ReadConfigValue(string key)
        {
            try
            {
                if (!File.Exists(ConfigFilePath)) return "";

                foreach (string line in File.ReadAllLines(ConfigFilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;

                    string k = line.Substring(0, eq).Trim();
                    if (string.Equals(k, key, StringComparison.OrdinalIgnoreCase))
                    {
                        return line.Substring(eq + 1).Trim();
                    }
                }
            }
            catch { }

            return "";
        }

        private string GetServerAddressForSave()
        {
            string current = _txtServer != null ? _txtServer.Text.Trim() : "";
            if (!string.IsNullOrEmpty(current)) return current;

            string saved = ReadConfigValue("server");
            if (!string.IsNullOrEmpty(saved)) return saved;

            return DefaultServerAddress;
        }

        private void LoadConfig()
        {
            _chkDemoMode.Checked = false;
            _txtServer.Text = DefaultServerAddress;
            _chkIntegratedSecurity.Checked = true;
            _cmbDatabases.Items.Clear();
            _cmbDatabases.Items.Add(DefaultDatabaseName);
            _cmbDatabases.SelectedIndex = 0;

            if (!File.Exists(ConfigFilePath))
            {
                // 최초 실행 기본값
                _chkDemoMode.Checked = false;
                _txtServer.Text = DefaultServerAddress;
                _chkIntegratedSecurity.Checked = true;
                _cmbDatabases.Items.Clear();
                _cmbDatabases.Items.Add(DefaultDatabaseName);
                _cmbDatabases.SelectedIndex = 0;
                return;
            }

            try
            {
                foreach (string line in File.ReadAllLines(ConfigFilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;
                    string k = line.Substring(0, eq).Trim();
                    string v = line.Substring(eq + 1).Trim();
                    int iv;
                    switch (k)
                    {
                        case "useMock":          _chkDemoMode.Checked = false; break;
                        case "server":           _txtServer.Text = string.IsNullOrEmpty(v) ? DefaultServerAddress : v; break;
                        case "integratedSec":    _chkIntegratedSecurity.Checked = true; break;
                        case "user":             _txtUser.Text = v; break;
                        case "database":
                            if (string.IsNullOrEmpty(v)) break;
                            _cmbDatabases.Items.Clear();
                            _cmbDatabases.Items.Add(v);
                            _cmbDatabases.SelectedIndex = 0;
                            break;
                        case "distChartResolver": if (int.TryParse(v, out iv)) _distChartResolver = Math.Max(100, Math.Min(1200, iv)); break;
                        case "distUser":          if (int.TryParse(v, out iv)) _distUser          = Math.Max(100, Math.Min(1400, iv)); break;
                        case "distCard":          if (int.TryParse(v, out iv)) _distCard          = Math.Max(100, Math.Min(1400, iv)); break;
                        case "distLabel":         if (int.TryParse(v, out iv)) _distLabel         = Math.Max(100, Math.Min(1400, iv)); break;
                        case "distRx":            if (int.TryParse(v, out iv)) _distRx            = Math.Max(100, Math.Min(1400, iv)); break;

                    }
                }

                if (string.IsNullOrEmpty(_txtServer.Text.Trim()))
                {
                    _txtServer.Text = DefaultServerAddress;
                }

                if (_cmbDatabases.SelectedItem == null || string.IsNullOrEmpty(_cmbDatabases.SelectedItem.ToString()))
                {
                    _cmbDatabases.Items.Clear();
                    _cmbDatabases.Items.Add(DefaultDatabaseName);
                    _cmbDatabases.SelectedIndex = 0;
                }

                _chkIntegratedSecurity.Checked = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("설정 파일 로드 오류: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Save Configuration to config file next to exe (portable)
        internal void SaveConfig()
        {
            try
            {
                string serverAddress = GetServerAddressForSave();
                if (_txtServer != null && string.IsNullOrEmpty(_txtServer.Text.Trim()))
                {
                    _txtServer.Text = serverAddress;
                }

                using (StreamWriter sw = new StreamWriter(ConfigFilePath))
                {
                    sw.WriteLine("useMock=False");
                    sw.WriteLine("server="           + serverAddress);
                    sw.WriteLine("integratedSec=True");
                    sw.WriteLine("user="             + _txtUser.Text.Trim());
                    sw.WriteLine("database="         + (_cmbDatabases.SelectedItem != null ? _cmbDatabases.SelectedItem.ToString() : ""));
                    sw.WriteLine("distChartResolver="+ _distChartResolver);
                    sw.WriteLine("distUser="         + _distUser);
                    sw.WriteLine("distCard="         + _distCard);
                    sw.WriteLine("distLabel="        + _distLabel);
                    sw.WriteLine("distRx="           + _distRx);
                }
            }
            catch { }
        }



        // Save Configuration to config file (connection settings + layout positions)
        private void BtnSaveConfig_Click(object sender, EventArgs e)
        {
            // Read current SplitterDistance from every visible SplitContainer
            if (_splitChartResolver != null && _splitChartResolver.Visible) _distChartResolver = _splitChartResolver.SplitterDistance;
            if (_splitUser          != null && _splitUser.Visible)          _distUser          = _splitUser.SplitterDistance;
            if (_splitCard          != null && _splitCard.Visible)          _distCard          = _splitCard.SplitterDistance;
            if (_splitLabel         != null && _splitLabel.Visible)         _distLabel         = _splitLabel.SplitterDistance;
            if (_splitRx            != null && _splitRx.Visible)            _distRx            = _splitRx.SplitterDistance;


            SaveConfig();
            ShowToast("설정과 레이아웃 위치가 저장되었습니다.", Color.FromArgb(99, 102, 241));
        }

        private void BtnSaveLayout_Click(object sender, EventArgs e)
        {
            BtnSaveConfig_Click(sender, e);
        }

        // Form Closing - only save col widths (splitter positions saved by 위치저장 button)
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }

        // Load DB List from Server
        private void BtnLoadDbs_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_txtServer.Text.Trim()))
            {
                MessageBox.Show("서버 주소를 먼저 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(true); // load from master DB fallback
            this.Cursor = Cursors.WaitCursor;

            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string query = "SELECT name FROM master.dbo.sysdatabases WHERE dbid > 4 ORDER BY name";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            _cmbDatabases.Items.Clear();
                            while (reader.Read())
                            {
                                _cmbDatabases.Items.Add(reader.GetString(0));
                            }
                        }
                    }
                }

                if (_cmbDatabases.Items.Count > 0)
                {
                    _cmbDatabases.SelectedIndex = 0;
                    ShowToast("DB 목록 로드 성공", ColorEmerald);
                }
                else
                {
                    MessageBox.Show("불러온 데이터베이스가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("데이터베이스 목록을 가져오지 못했습니다:\n" + ex.Message, "연결 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        internal string BuildConnectionString(bool useMasterFallback)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            
            // Handle server and custom port if entered as IP,port
            builder.DataSource = _txtServer.Text.Trim();
            
            if (useMasterFallback || _cmbDatabases.SelectedItem == null)
            {
                builder.InitialCatalog = "master";
            }
            else
            {
                builder.InitialCatalog = _cmbDatabases.SelectedItem.ToString();
            }

            if (_chkIntegratedSecurity.Checked)
            {
                builder.IntegratedSecurity = true;
            }
            else
            {
                builder.UserID = _txtUser.Text.Trim();
                builder.Password = _txtPassword.Text;
            }

            // Trust certificate and set timeout to prevent hangs
            builder.TrustServerCertificate = true;
            builder.ConnectTimeout = 10;

            return builder.ConnectionString;
        }

        private void BtnSearchRx_Click(object sender, EventArgs e)
        {
            string name = _txtSearchName.Text.Trim();
            string jumin = _txtSearchJumin.Text.Replace("-", "").Trim();

            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(jumin))
            {
                MessageBox.Show("검색 조건(이름 또는 주민번호 앞7자리)을 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                SearchRxMock(name, jumin);
            }
            else
            {
                SearchRxProduction(name, jumin);
            }
        }

        // Demo search logic
        private void SearchRxMock(string name, string jumin)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("차트번호");
            dt.Columns.Add("환자 이름");
            dt.Columns.Add("주민번호");
            dt.Columns.Add("처방 횟수", typeof(int));
            dt.Columns.Add("최종 처방일");

            // Temporary aggregation dictionary
            var aggregated = new Dictionary<string, AggregatedResult>();

            foreach (var rx in _mockRxList)
            {
                // Normalize Jumin digits
                string rxJuminClean = rx.PatJuminNo.Replace("-", "");

                bool isMatch = false;
                if (!string.IsNullOrEmpty(jumin))
                {
                    // Prefix matching
                    isMatch = rxJuminClean.StartsWith(jumin);
                }
                else if (!string.IsNullOrEmpty(name))
                {
                    // Exact name matching
                    isMatch = rx.PatNm == name;
                }

                if (isMatch)
                {
                    if (!aggregated.ContainsKey(rx.ChrtNo))
                    {
                        aggregated[rx.ChrtNo] = new AggregatedResult { Name = rx.PatNm, Jumin = rx.PatJuminNo, Count = 0, LastDate = rx.MedYmd };
                    }

                    var current = aggregated[rx.ChrtNo];
                    int newCount = current.Count + 1;
                    string newLastDate = current.LastDate;

                    if (DateTime.Parse(rx.MedYmd) > DateTime.Parse(current.LastDate))
                    {
                        newLastDate = rx.MedYmd;
                    }

                    current.Count = newCount;
                    current.LastDate = newLastDate;
                }
            }

            foreach (var kvp in aggregated)
            {
                string formattedJumin = FormatJuminPrefix(kvp.Value.Jumin);
                dt.Rows.Add(kvp.Key, kvp.Value.Name, formattedJumin, kvp.Value.Count, kvp.Value.LastDate);
            }

            _dgvResults.DataSource = dt;
            
            // Adjust grid column widths
            if (_dgvResults.Columns.Count > 0)
            {
                _dgvResults.Columns[0].Width = 90;
                _dgvResults.Columns[1].Width = 90;
                _dgvResults.Columns[2].Width = 140;
                _dgvResults.Columns[3].Width = 80;
                _dgvResults.Columns[4].Width = 120;
            }

            ShowToast(string.Format("검색 완료: {0}건 조회", dt.Rows.Count), ColorEmerald);
        }

        internal string GetDateColumn(SqlConnection conn)
        {
            if (_detectedDateColumn != null) return _detectedDateColumn;
            _detectedDateColumn = ""; 

            try
            {
                string query = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'tbsid040_03'";
                List<string> columns = new List<string>();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            columns.Add(reader.GetString(0).ToLower());
                        }
                    }
                }

                string[] candidates = { 
                    "med_ymd", "medymd", "med_date", "meddate", 
                    "presc_date", "prescymd", "presc_ymd", 
                    "med_dt", "meddt", "med_ym", "rx_date", "rxdate" 
                };

                foreach (var candidate in candidates)
                {
                    if (columns.Contains(candidate))
                    {
                        _detectedDateColumn = candidate;
                        return _detectedDateColumn;
                    }
                }

                foreach (var col in columns)
                {
                    if (col.Contains("date") || col.Contains("ymd") || col.Contains("dt"))
                    {
                        _detectedDateColumn = col;
                        return _detectedDateColumn;
                    }
                }
            }
            catch (Exception)
            {
                // Fail silently and fallback to empty
            }

            return _detectedDateColumn;
        }

        // Live Database Search
        private void SearchRxProduction(string name, string jumin)
        {
            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            try
            {
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    string dateCol = GetDateColumn(conn);
                    string selectDate = "";
                    if (!string.IsNullOrEmpty(dateCol))
                    {
                        selectDate = ", MAX(" + dateCol + ") as [최종 처방일]";
                    }
                    else
                    {
                        selectDate = ", NULL as [최종 처방일]";
                    }

                    string sql = @"
                        SELECT chrtno as [차트번호], 
                               pat_nm as [환자 이름], 
                               pat_jumin_no as [주민번호], 
                               COUNT(*) as [처방 횟수]" + selectDate + @"
                        FROM tbsid040_03
                        WHERE 1=1";

                    if (!string.IsNullOrEmpty(jumin))
                    {
                        sql += " AND REPLACE(pat_jumin_no, '-', '') LIKE @jumin + '%'";
                        if (!string.IsNullOrEmpty(name))
                        {
                            if (name.Length >= 3)
                            {
                                sql += " AND pat_nm LIKE @name_prefix + '%'";
                            }
                            else
                            {
                                sql += " AND pat_nm LIKE @name_first + '%'";
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(name))
                    {
                        sql += " AND pat_nm = @name";
                    }

                    sql += " GROUP BY chrtno, pat_nm, pat_jumin_no";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        if (!string.IsNullOrEmpty(jumin))
                        {
                            cmd.Parameters.AddWithValue("@jumin", jumin);
                            if (!string.IsNullOrEmpty(name))
                            {
                                if (name.Length >= 3)
                                {
                                    cmd.Parameters.AddWithValue("@name_prefix", name.Substring(0, 2));
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@name_first", name.Substring(0, 1));
                                }
                            }
                        }
                        else if (!string.IsNullOrEmpty(name))
                        {
                            cmd.Parameters.AddWithValue("@name", name);
                        }

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }

                // Format Jumin in data table
                foreach (DataRow row in dt.Rows)
                {
                    row["주민번호"] = FormatJuminPrefix(row["주민번호"].ToString());
                    
                    // Format Date if datetime
                    if (row["최종 처방일"] != DBNull.Value)
                    {
                        string rawDate = row["최종 처방일"].ToString();
                        // Simplify datetime strings
                        if (rawDate.Contains(" "))
                        {
                            row["최종 처방일"] = rawDate.Split(' ')[0];
                        }
                    }
                }

                _dgvResults.DataSource = dt;
                
                if (_dgvResults.Columns.Count > 0)
                {
                    _dgvResults.Columns[0].Width = 90;
                    _dgvResults.Columns[1].Width = 90;
                    _dgvResults.Columns[2].Width = 140;
                    _dgvResults.Columns[3].Width = 80;
                    _dgvResults.Columns[4].Width = 120;
                }

                ShowToast(string.Format("검색 완료: {0}건 조회", dt.Rows.Count), ColorEmerald);
            }
            catch (Exception ex)
            {
                MessageBox.Show("처방 테이블 검색 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        // Grid format event to highlight rows with name discrepancy
        private void DgvResults_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string searchNameVal = _txtSearchName.Text.Trim();
            if (string.IsNullOrEmpty(searchNameVal)) return;

            // Col 1 is "환자 이름"
            var cellName = _dgvResults.Rows[e.RowIndex].Cells[1].Value;
            if (cellName != null && cellName.ToString() != searchNameVal)
            {
                // Discrepancy! Highlight background red
                _dgvResults.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(127, 29, 29); // Dark Red
                _dgvResults.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
            }
        }

        // Grid Click -> Copy and Query details automatically
        private void DgvResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var chartVal = _dgvResults.Rows[e.RowIndex].Cells[0].Value;
            if (chartVal == null) return;

            string chrtno = chartVal.ToString().Trim();
            
            // 1. Copy to Clipboard
            try
            {
                Clipboard.SetText(chrtno);
            }
            catch (Exception)
            {
                // Clipboard access might fail under rare conditions, fail silently
            }

            // 2. Set search box and run lookup
            _txtSearchChrtNo.Text = chrtno;
            SearchCustomer(chrtno);

            // 3. Show Toast Notice
            ShowToast(string.Format("차트번호 {0} 복사 및 조회 완료!", chrtno), ColorEmerald);
        }

        private void BtnSearchCust_Click(object sender, EventArgs e)
        {
            string chrtno = _txtSearchChrtNo.Text.Trim();
            if (string.IsNullOrEmpty(chrtno))
            {
                MessageBox.Show("조회할 차트번호를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SearchCustomer(chrtno);
        }

        // Master Customer Search logic
        private void SearchCustomer(string chrtno)
        {
            if (_chkDemoMode.Checked)
            {
                SearchCustMock(chrtno);
            }
            else
            {
                SearchCustProduction(chrtno);
            }
        }

        private void SearchCustMock(string chrtno)
        {
            var cust = _mockCustList.Find(c => c.ChrtNo == chrtno);
            if (cust == null)
            {
                ClearCustomerDetails("해당 고객 정보가 존재하지 않습니다.");
                return;
            }

            // Load Customer detail fields
            _lblCustNameTitle.Text = cust.PatNm;
            _lblCustChrtNoBadge.Text = cust.ChrtNo;
            _lblCustChrtNo.Text = cust.ChrtNo;
            _lblCustName.Text = cust.PatNm;
            _lblCustJumin.Text = FormatJuminFull(cust.PatJuminNo);
            _lblCustPhone.Text = cust.Phone;
            _lblCustAddress.Text = cust.Address;
            _lblCustFirstVisit.Text = cust.FirstVisit;

            // Load Recent Rx list
            _lstRxHistory.Items.Clear();
            foreach (var rx in _mockRxList)
            {
                if (rx.ChrtNo == chrtno)
                {
                    _lstRxHistory.Items.Add(string.Format("[{0}] {1}", rx.MedYmd, rx.Medicine));
                }
            }
        }

        private void SearchCustProduction(string chrtno)
        {
            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    // Query customer master
                    string queryCust = "SELECT chrtno, pat_nm, jumin_no FROM tbsit000_01 WHERE chrtno = @chrtno";
                    bool found = false;

                    using (SqlCommand cmd = new SqlCommand(queryCust, conn))
                    {
                        cmd.Parameters.AddWithValue("@chrtno", chrtno);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                found = true;
                                string name = reader["pat_nm"].ToString();
                                string chrt = reader["chrtno"].ToString();
                                _lblCustNameTitle.Text = name;
                                _lblCustChrtNoBadge.Text = chrt;
                                _lblCustChrtNo.Text = chrt;
                                _lblCustName.Text = name;
                                _lblCustJumin.Text = FormatJuminFull(reader["jumin_no"].ToString());
                                _lblCustPhone.Text = "제공되지 않음 (컬럼 없음)";
                                _lblCustAddress.Text = "제공되지 않음 (컬럼 없음)";
                                _lblCustFirstVisit.Text = "-";
                            }
                        }
                    }

                    if (!found)
                    {
                        ClearCustomerDetails("해당 고객 정보가 존재하지 않습니다.");
                        return;
                    }

                    // Query Rx History
                    string dateCol = GetDateColumn(conn);
                    _lstRxHistory.Items.Clear();

                    if (!string.IsNullOrEmpty(dateCol))
                    {
                        string queryRx = "SELECT TOP 10 " + dateCol + " FROM tbsid040_03 WHERE chrtno = @chrtno ORDER BY " + dateCol + " DESC";
                        using (SqlCommand cmdRx = new SqlCommand(queryRx, conn))
                        {
                            cmdRx.Parameters.AddWithValue("@chrtno", chrtno);
                            using (SqlDataReader readerRx = cmdRx.ExecuteReader())
                            {
                                while (readerRx.Read())
                                {
                                    string dateStr = readerRx[0].ToString();
                                    if (dateStr.Contains(" ")) dateStr = dateStr.Split(' ')[0];
                                    
                                    _lstRxHistory.Items.Add(string.Format("[{0}] 처방 기록 존재", dateStr));
                                }
                            }
                        }
                    }
                    else
                    {
                        string queryCount = "SELECT COUNT(*) FROM tbsid040_03 WHERE chrtno = @chrtno";
                        using (SqlCommand cmdCount = new SqlCommand(queryCount, conn))
                        {
                            cmdCount.Parameters.AddWithValue("@chrtno", chrtno);
                            int rxCount = (int)cmdCount.ExecuteScalar();
                            if (rxCount > 0)
                            {
                                _lstRxHistory.Items.Add(string.Format("총 {0}건의 처방 기록이 존재합니다. (날짜 컬럼 없음)", rxCount));
                            }
                        }
                    }

                    if (_lstRxHistory.Items.Count == 0)
                    {
                        _lstRxHistory.Items.Add("최근 처방 기록이 없습니다.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("고객 정보 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ClearCustomerDetails(string message)
        {
            _lblCustChrtNoBadge.Text = "-";
            _lblCustNameTitle.Text = message;
            _lblCustChrtNo.Text = "-";
            _lblCustName.Text = message;
            _lblCustJumin.Text = "-";
            _lblCustPhone.Text = "-";
            _lblCustAddress.Text = "-";
            _lblCustFirstVisit.Text = "-";
            _lstRxHistory.Items.Clear();
        }

        private void BtnSyncCharts_Click(object sender, EventArgs e)
        {
            DialogResult dr1 = MessageBox.Show(
                "처방 테이블(tbsid040_03)의 차트번호를 활성화된 고객 마스터(tbsit000_01)의 차트번호와 강제로 일치시킵니다.\n이 작업은 데이터베이스를 직접 수정합니다. 진행하시겠습니까?",
                "차트번호 동기화 경고 (1/2)",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr1 != DialogResult.Yes) return;

            DialogResult dr2 = MessageBox.Show(
                "정말로 처방-고객 차트번호를 동기화하시겠습니까? 실행 후에는 되돌릴 수 없습니다.",
                "동기화 최종 확인 (2/2)",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr2 != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                SyncChartsMock();
            }
            else
            {
                SyncChartsProduction();
            }
        }

        private void SyncChartsMock()
        {
            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                System.Threading.Thread.Sleep(800); // simulate delay

                int updatedRows = 0;
                foreach (var rx in _mockRxList)
                {
                    var activeCusts = _mockCustList.FindAll(c => 
                        c.JuminEncrypt == rx.JuminEncrypt && 
                        c.CusAct == "1" && 
                        !string.IsNullOrEmpty(c.JuminNo) && 
                        char.IsDigit(c.JuminNo[0])
                    );

                    if (activeCusts.Count == 1)
                    {
                        var activeCust = activeCusts[0];
                        if (activeCust.ChrtNo != rx.ChrtNo || activeCust.PatNm != rx.PatNm)
                        {
                            rx.ChrtNo = activeCust.ChrtNo;
                            rx.PatNm = activeCust.PatNm;
                            updatedRows++;
                        }
                    }
                }

                this.BeginInvoke((Action)(() =>
                {
                    this.Cursor = Cursors.Default;

                    ShowToast(string.Format("[데모] {0}건 동기화 완료", updatedRows), ColorEmerald);
                    MessageBox.Show(string.Format("[데모] 차트번호 동기화가 완료되었습니다.\n\n- 변경 완료: {0}건", updatedRows), "동기화 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (!string.IsNullOrEmpty(_txtSearchName.Text) || !string.IsNullOrEmpty(_txtSearchJumin.Text))
                    {
                        BtnSearchRx_Click(null, null);
                    }
                }));
            });
        }

        private void SyncChartsProduction()
        {
            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            
            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                int affectedRows = 0;
                string errorMsg = null;

                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();

                        string sql = @"
                            UPDATE B 
                            SET B.CHRTNO = A.CHRTNO,
                                B.PAT_NM = A.PAT_NM
                            FROM TBSIT000_01 AS A
                            INNER JOIN TBSID040_03 AS B ON A.JUMIN_ENCRYPT = B.JUMIN_ENCRYPT
                            WHERE (B.CHRTNO <> A.CHRTNO OR B.PAT_NM <> A.PAT_NM)
                            AND A.CUSACT = '1'
                            AND A.JUMIN_NO NOT LIKE '%*%'
                            AND A.JUMIN_NO <> ''
                            AND A.JUMIN_ENCRYPT IN (
                                SELECT JUMIN_ENCRYPT 
                                FROM TBSIT000_01 
                                WHERE CUSACT = '1' 
                                  AND JUMIN_NO NOT LIKE '%*%'
                                  AND JUMIN_NO <> ''
                                GROUP BY JUMIN_ENCRYPT 
                                HAVING COUNT(DISTINCT CHRTNO) = 1
                            )";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            affectedRows = cmd.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorMsg = ex.Message;
                }

                this.BeginInvoke((Action)(() =>
                {
                    this.Cursor = Cursors.Default;

                    if (errorMsg != null)
                    {
                        MessageBox.Show("동기화 실행 중 오류 발생:\n" + errorMsg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        ShowToast(string.Format("동기화 완료: {0}건 변경됨", affectedRows), ColorEmerald);
                        MessageBox.Show(string.Format("차트번호 동기화가 성공적으로 완료되었습니다.\n\n- 변경 완료: {0}건", affectedRows), "동기화 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        if (!string.IsNullOrEmpty(_txtSearchName.Text) || !string.IsNullOrEmpty(_txtSearchJumin.Text))
                        {
                            BtnSearchRx_Click(null, null);
                        }
                    }
                }));
            });
        }

        // Toast message display
        private void ShowToast(string message, Color color)
        {
            _lblToast.Text = message;
            _lblToast.ForeColor = color;
            _lblToast.Visible = true;

            _toastTimer.Stop();
            _toastTimer.Start();
            UpdateStatus(message);
        }

        // Helper: format Jumin to prefix format (571029-2******)
        private string FormatJuminPrefix(string jumin)
        {
            if (string.IsNullOrEmpty(jumin)) return "";
            string clean = Regex.Replace(jumin, "[^0-9]", "");
            if (clean.Length >= 7)
            {
                return clean.Substring(0, 6) + "-" + clean.Substring(6, 1) + "******";
            }
            else if (clean.Length == 6)
            {
                return clean.Substring(0, 6) + "-";
            }
            return jumin;
        }

        // Helper: format full Jumin (mask back 6 digits)
        internal string FormatJuminFull(string jumin)
        {
            if (string.IsNullOrEmpty(jumin)) return "";
            string clean = Regex.Replace(jumin, "[^0-9]", "");
            if (clean.Length >= 7)
            {
                return clean.Substring(0, 6) + "-" + clean.Substring(6, 1) + "******";
            }
            return jumin;
        }

        private void BtnRestoreCust_Click(object sender, EventArgs e)
        {
            string chrtno = _txtSearchChrtNo.Text.Trim();
            if (string.IsNullOrEmpty(chrtno))
            {
                MessageBox.Show("복구할 차트번호를 먼저 입력하거나 조회해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get restoration candidates from prescriptions
            List<PatientGroup> candidates = GetRestoreCandidates(chrtno);
            if (candidates.Count == 0)
            {
                MessageBox.Show("해당 차트번호로 등록된 처방 내역이 없어 복구 기준 정보를 찾을 수 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (candidates.Count == 1)
            {
                var pg = candidates[0];
                DialogResult dr = MessageBox.Show(
                    string.Format("차트번호 [{0}]의 고객 마스터 정보를 처방 내역의 [{1}] 님 정보로 복구하시겠습니까?\n\n- 복구할 이름: {1}\n- 복구할 주민번호: {2}\n- 세대주: {3}",
                        chrtno, pg.Name, pg.Jumin, pg.FamNm),
                    "고객 정보 복구 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (dr == DialogResult.Yes)
                {
                    RestoreMaster(chrtno, pg);
                    SearchCustomer(chrtno); // Refresh main window info
                }
            }
            else
            {
                // Multiple candidates, show selection form
                using (var form = new RestoreSelectionForm(candidates))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        var pg = new PatientGroup
                        {
                            Name = form.SelectedName,
                            Jumin = form.SelectedJumin,
                            JuminEncrypt = form.SelectedEncrypt,
                            FamNm = form.SelectedFamNm
                        };
                        RestoreMaster(chrtno, pg);
                        SearchCustomer(chrtno); // Refresh main window info
                    }
                }
            }
        }

        public void RestoreMaster(string chrtno, PatientGroup pg)
        {
            if (_chkDemoMode.Checked)
            {
                RestoreMasterDemo(chrtno, pg);
            }
            else
            {
                RestoreMasterProduction(chrtno, pg);
            }
        }

        private void RestoreMasterDemo(string chrtno, PatientGroup pg)
        {
            var cust = _mockCustList.Find(c => c.ChrtNo == chrtno);
            if (cust != null)
            {
                cust.PatNm = pg.Name;
                cust.PatJuminNo = pg.Jumin;
                cust.JuminNo = pg.Jumin.Replace("-", "");
                cust.JuminEncrypt = pg.JuminEncrypt;
                cust.FamNm = pg.FamNm;

                MessageBox.Show("고객 마스터 정보가 성공적으로 복구되었습니다. (데모 모드)", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                var newCust = new MockCust
                {
                    ChrtNo = chrtno,
                    PatNm = pg.Name,
                    PatJuminNo = pg.Jumin,
                    JuminNo = pg.Jumin.Replace("-", ""),
                    JuminEncrypt = pg.JuminEncrypt,
                    FamNm = pg.FamNm,
                    Phone = "010-0000-0000",
                    Address = "임시 등록 주소",
                    FirstVisit = DateTime.Now.ToString("yyyy-MM-dd"),
                    CusAct = "1"
                };
                _mockCustList.Add(newCust);
                MessageBox.Show("고객 마스터에 새롭게 정보가 복구/등록되었습니다. (데모 모드)", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        internal bool TableExists(SqlConnection conn, string tableName, SqlTransaction trans = null)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM sys.tables WHERE name = @tableName", conn))
                {
                    if (trans != null) cmd.Transaction = trans;
                    cmd.Parameters.AddWithValue("@tableName", tableName);
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private void RestoreMasterProduction(string chrtno, PatientGroup pg)
        {
            string connStr = BuildConnectionString(false);
            try
            {
                string targetJumin = pg.Jumin;
                string targetEncrypt = pg.JuminEncrypt;
                string targetFamNm = pg.FamNm;

                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    if (TableExists(conn, "TEMP_MAPPING_CHRTNO"))
                    {
                        string tempSql = "SELECT JUMIN_NO, JUMIN_ENCRYPT FROM TEMP_MAPPING_CHRTNO WHERE chrtno = @chrtno AND pat_nm = @pat_nm";
                        using (SqlCommand tempCmd = new SqlCommand(tempSql, conn))
                        {
                            tempCmd.Parameters.AddWithValue("@chrtno", chrtno);
                            tempCmd.Parameters.AddWithValue("@pat_nm", pg.Name);
                            using (SqlDataReader r = tempCmd.ExecuteReader())
                            {
                                if (r.Read())
                                {
                                    targetJumin = r["JUMIN_NO"].ToString();
                                    targetEncrypt = r["JUMIN_ENCRYPT"].ToString();
                                }
                            }
                        }
                    }

                    string updateSql = @"
                        UPDATE tbsit000_01
                        SET pat_nm = @pat_nm,
                            jumin_no = @jumin_no,
                            jumin_encrypt = @jumin_encrypt,
                            fam_nm = @fam_nm,
                            proc_dtime = @proc_dtime
                        WHERE chrtno = @chrtno";

                    using (SqlCommand cmd = new SqlCommand(updateSql, conn))
                    {
                        cmd.Parameters.AddWithValue("@pat_nm", pg.Name);
                        cmd.Parameters.AddWithValue("@jumin_no", targetJumin);
                        cmd.Parameters.AddWithValue("@jumin_encrypt", targetEncrypt);
                        cmd.Parameters.AddWithValue("@fam_nm", targetFamNm);
                        cmd.Parameters.AddWithValue("@proc_dtime", DateTime.Now.ToString("yyyyMMddHHmmss"));
                        cmd.Parameters.AddWithValue("@chrtno", chrtno);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                        {
                            MessageBox.Show("고객 마스터 정보가 성공적으로 복구되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string insertSql = @"
                                INSERT INTO tbsit000_01 (chrtno, pat_seq, pat_nm, jumin_no, jumin_encrypt, fam_nm, cusact, proc_dtime)
                                VALUES (@chrtno, 1, @pat_nm, @jumin_no, @jumin_encrypt, @fam_nm, '1', @proc_dtime)";

                            using (SqlCommand insCmd = new SqlCommand(insertSql, conn))
                            {
                                insCmd.Parameters.AddWithValue("@chrtno", chrtno);
                                insCmd.Parameters.AddWithValue("@pat_nm", pg.Name);
                                insCmd.Parameters.AddWithValue("@jumin_no", targetJumin);
                                insCmd.Parameters.AddWithValue("@jumin_encrypt", targetEncrypt);
                                insCmd.Parameters.AddWithValue("@fam_nm", targetFamNm);
                                insCmd.Parameters.AddWithValue("@proc_dtime", DateTime.Now.ToString("yyyyMMddHHmmss"));
                                insCmd.ExecuteNonQuery();
                            }
                            MessageBox.Show("고객 마스터에 새롭게 정보가 복구/등록되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("고객 정보 복구 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal List<PatientGroup> GetRestoreCandidates(string chrtno)
        {
            var candidates = new List<PatientGroup>();

            if (_chkDemoMode.Checked)
            {
                var dict = new Dictionary<string, PatientGroup>();
                foreach (var rx in _mockRxList)
                {
                    if (rx.ChrtNo == chrtno)
                    {
                        string key = rx.PatNm + "|" + rx.PatJuminNo;
                        if (!dict.ContainsKey(key))
                        {
                            dict[key] = new PatientGroup 
                            { 
                                Name = rx.PatNm, 
                                Jumin = rx.PatJuminNo, 
                                Count = 0, 
                                JuminEncrypt = rx.JuminEncrypt, 
                                FamNm = rx.PatNm == "천미선" ? "백승현" : "임광묵" 
                            };
                        }
                        dict[key].Count++;
                    }
                }
                candidates.AddRange(dict.Values);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string queryRx = @"
                            SELECT pat_nm, pat_jumin_no, COUNT(*) as rx_count, MIN(jumin_encrypt) as jumin_encrypt, MIN(fam_nm) as fam_nm
                            FROM tbsid040_03
                            WHERE chrtno = @chrtno
                            GROUP BY pat_nm, pat_jumin_no";

                        using (SqlCommand cmd = new SqlCommand(queryRx, conn))
                        {
                            cmd.Parameters.AddWithValue("@chrtno", chrtno);
                            using (SqlDataReader r = cmd.ExecuteReader())
                            {
                                while (r.Read())
                                {
                                    candidates.Add(new PatientGroup
                                    {
                                        Name = r["pat_nm"].ToString(),
                                        Jumin = r["pat_jumin_no"].ToString(),
                                        Count = Convert.ToInt32(r["rx_count"]),
                                        JuminEncrypt = r["jumin_encrypt"].ToString(),
                                        FamNm = r["fam_nm"].ToString()
                                    });
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("처방 내역 정보 조회 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return candidates;
        }

        private void BtnSearchDoctors_Click(object sender, EventArgs e)
        {
            _btnSearchDoctors.Enabled = false;
            _btnSearchDoctors.Text = "⏳ 조회 중...";
            LoadDoctorGrid();
        }

        private void BtnDeleteDoctors_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(
                "중복 등록된 의사면허 데이터를 일괄 정리하시겠습니까?\n" +
                "(이 작업은 요양기관, 의사면허번호, 근무지번호 기준 중복된 데이터 중 가장 최근에 생성된(SEQ가 가장 큰) 데이터 1건만 남기고 나머지를 삭제하며, 실행 전 백업 테이블(tbsim000_12_back20260701)을 생성합니다.)",
                "의사면허 일괄 삭제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            _btnDeleteDoctors.Enabled = false;
            _btnDeleteDoctors.Text = "⏳ 정리 중...";
            DeleteDuplicateDoctors();
        }

        private void LoadDoctorGrid()
        {
            if (_chkDemoMode.Checked)
            {
                ScanDoctorsDemo();
            }
            else
            {
                ScanDoctorsProduction();
            }
        }

        private void ScanDoctorsDemo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("일련번호(SEQ)", typeof(int));
            dt.Columns.Add("요양기관기호");
            dt.Columns.Add("요양기관명");
            dt.Columns.Add("의사면허번호(DC_ID)");
            dt.Columns.Add("의사명");
            dt.Columns.Add("의사구분");

            var dupList = _mockDoctorList
                .GroupBy(d => new { d.Ykiho, d.DcId })
                .Where(g => g.Count() > 1)
                .SelectMany(g => g)
                .OrderBy(d => d.Ykiho)
                .ThenBy(d => d.DcId)
                .ThenBy(d => d.Seq)
                .ToList();

            foreach (var d in dupList)
            {
                dt.Rows.Add(d.Seq, d.Ykiho, d.YoyangNm, d.DcId, d.DcName, d.DrGubun);
            }

            _dgvDoctors.DataSource = dt;
            if (_dgvDoctors.Columns.Count > 0)
            {
                _dgvDoctors.Columns[0].Width = 110;
                _dgvDoctors.Columns[1].Width = 120;
                _dgvDoctors.Columns[2].Width = 150;
                _dgvDoctors.Columns[3].Width = 150;
                _dgvDoctors.Columns[4].Width = 120;
                _dgvDoctors.Columns[5].Width = 120;
            }

            _btnSearchDoctors.Enabled = true;
            _btnSearchDoctors.Text = "🔍 중복 의사면허 조회";
            _lblDoctorStatus.Text = string.Format("중복 검출: {0}건", dupList.Count);
        }

        private void ScanDoctorsProduction()
        {
            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT T1.SEQ as [일련번호(SEQ)],
                                   T1.YKIHO as [요양기관기호],
                                   T1.YOYANG_NM as [요양기관명],
                                   T1.DC_ID as [의사면허번호(DC_ID)],
                                   T1.DC_NAME as [의사명],
                                   T1.DR_GUBUN as [의사구분]
                            FROM TBSIM000_12 AS T1
                            INNER JOIN (
                                 SELECT *
                                 FROM (SELECT SEQ, YKIHO, DC_ID, ROW_NUMBER() OVER(PARTITION BY YKIHO, DC_ID ORDER BY SEQ, YKIHO) AS ROW_NUM 
                                       FROM TBSIM000_12) AS A1
                                 WHERE ROW_NUM > 1) AS T2
                            ON T1.DC_ID = T2.DC_ID
                            AND T1.YKIHO = T2.YKIHO
                            ORDER BY T1.YKIHO ASC, T1.DC_ID ASC, T1.SEQ ASC";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.CommandTimeout = 300;
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }

                    this.BeginInvoke((Action)(() =>
                    {
                        _dgvDoctors.DataSource = dt;
                        if (_dgvDoctors.Columns.Count > 0)
                        {
                            _dgvDoctors.Columns[0].Width = 110;
                            _dgvDoctors.Columns[1].Width = 120;
                            _dgvDoctors.Columns[2].Width = 150;
                            _dgvDoctors.Columns[3].Width = 150;
                            _dgvDoctors.Columns[4].Width = 120;
                            _dgvDoctors.Columns[5].Width = 120;
                        }
                        _btnSearchDoctors.Enabled = true;
                        _btnSearchDoctors.Text = "🔍 중복 의사면허 조회";
                        _lblDoctorStatus.Text = string.Format("중복 검출: {0}건", dt.Rows.Count);
                        this.Cursor = Cursors.Default;
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        _btnSearchDoctors.Enabled = true;
                        _btnSearchDoctors.Text = "🔍 중복 의사면허 조회";
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("의사면허 조회 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void DeleteDuplicateDoctors()
        {
            if (_chkDemoMode.Checked)
            {
                DeleteDoctorsDemo();
            }
            else
            {
                DeleteDoctorsProduction();
            }
        }

        private void DeleteDoctorsDemo()
        {
            int deletedCount = 0;
            var groups = _mockDoctorList.GroupBy(d => new { d.Ykiho, d.DcId }).ToList();
            foreach (var g in groups)
            {
                if (g.Count() > 1)
                {
                    var sorted = g.OrderByDescending(d => d.Seq).ToList();
                    for (int i = 1; i < sorted.Count; i++)
                    {
                        _mockDoctorList.Remove(sorted[i]);
                        deletedCount++;
                    }
                }
            }

            _btnDeleteDoctors.Enabled = true;
            _btnDeleteDoctors.Text = "🗑️ 중복 의사면허 일괄 삭제";
            MessageBox.Show(string.Format("[데모] 중복 데이터 {0}건이 정상적으로 정리(삭제)되었습니다. (최신 데이터 보존)", deletedCount), "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ScanDoctorsDemo();
        }

        private void DeleteDoctorsProduction()
        {
            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int affectedRows = 0;
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        
                        // 1. 현재 처방의사 발급기관 데이터 백업
                        string backupSql = @"
                            IF OBJECT_ID('pm_main..tbsim000_12_back20260701', 'U') IS NOT NULL
                                DROP TABLE pm_main..tbsim000_12_back20260701;

                            SELECT * INTO pm_main..tbsim000_12_back20260701 FROM tbsim000_12;";

                        using (SqlCommand backupCmd = new SqlCommand(backupSql, conn))
                        {
                            backupCmd.CommandTimeout = 300;
                            backupCmd.ExecuteNonQuery();
                        }

                        // 2. 중복된 데이터 최신데이터만 남기고 삭제하기
                        string deleteSql = @"
                            DELETE a 
                            FROM pm_main..tbsim000_12 a 
                            WHERE NOT EXISTS (
                                SELECT b.* 
                                FROM (
                                    SELECT MAX(seq) AS seq, ykiho, dc_id, wrk_no 
                                    FROM pm_main..tbsim000_12
                                    GROUP BY ykiho, dc_id, wrk_no
                                ) b
                                WHERE a.seq = b.seq
                            );";

                        using (SqlCommand deleteCmd = new SqlCommand(deleteSql, conn))
                        {
                            deleteCmd.CommandTimeout = 300;
                            affectedRows = deleteCmd.ExecuteNonQuery();
                        }
                    }

                    this.BeginInvoke((Action)(() =>
                    {
                        _btnDeleteDoctors.Enabled = true;
                        _btnDeleteDoctors.Text = "🗑️ 중복 의사면허 일괄 삭제";
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(string.Format("중복 데이터 {0}건이 정상적으로 정리(삭제)되었습니다.", affectedRows), "정리 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ScanDoctorsProduction();
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        _btnDeleteDoctors.Enabled = true;
                        _btnDeleteDoctors.Text = "🗑️ 중복 의사면허 일괄 삭제";
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("의사면허 일괄 삭제 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

private void BtnRunDbRecovery_Click(object sender, EventArgs e)
        {
            List<string> selectedDbs = new List<string>();
            if (_chkDbPmMain.Checked) selectedDbs.Add("pm_main");
            if (_chkDbPmplusDums.Checked) selectedDbs.Add("pmplus_dums");
            if (_chkDbPmplusImage.Checked) selectedDbs.Add("pmplus_image");
            if (_chkDbPmplusJoblog.Checked) selectedDbs.Add("pmplus_joblog");

            if (selectedDbs.Count == 0)
            {
                MessageBox.Show("복구 작업을 진행할 데이터베이스를 최소 1개 이상 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(
                "🚨 [경고] 정말로 선택한 데이터베이스에 대해 응급 복구 작업을 실행하시겠습니까?\n\n" +
                "- 대상 DB: " + string.Join(", ", selectedDbs.ToArray()) + "\n\n" +
                "※ 중요 안내:\n" +
                "1. 복구 옵션(REPAIR_ALLOW_DATA_LOSS)에 의해 일부 손상된 데이터가 영구적으로 손실(삭제)될 수 있습니다.\n" +
                "2. 실행 중 데이터베이스가 싱글 유저(SINGLE_USER) 모드로 강제 제어되므로 조제/청구 업무가 일시 차단됩니다.\n" +
                "3. 복구가 진행되는 수 분 동안 다른 약국 프로그램 접속이 끊어질 수 있습니다.",
                "데이터베이스 최종 응급 복구 경고",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Stop
            );

            if (dr != DialogResult.Yes) return;

            DialogResult dr2 = MessageBox.Show(
                "동의하십니까? 데이터가 삭제될 위험이 있으며 이에 대한 백업은 별도로 책임져야 합니다.\n" +
                "계속 진행하시려면 [예]를 누르십시오.",
                "데이터 손실 위험 동의 서명",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr2 != DialogResult.Yes) return;

            _btnRunDbRecovery.Enabled = false;
            _btnRunDbRecovery.Text = "⏳ 복구 진행 중...";
            _txtDbRecoveryLog.Text = "";
            AppendRecoveryLog("============================================================\r\n");
            AppendRecoveryLog(string.Format("▶ [{0}] DB 응급 복구 및 무결성 검사 시퀀스 가동 시작\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendRecoveryLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                RunDbRecoveryDemo(selectedDbs);
            }
            else
            {
                RunDbRecoveryProduction(selectedDbs);
            }
        }

        private void BtnShrinkDb_Click(object sender, EventArgs e)
        {
            List<string> selectedDbs = new List<string>();
            if (_chkDbPmMain.Checked) selectedDbs.Add("pm_main");
            if (_chkDbPmplusDums.Checked) selectedDbs.Add("pmplus_dums");
            if (_chkDbPmplusImage.Checked) selectedDbs.Add("pmplus_image");
            if (_chkDbPmplusJoblog.Checked) selectedDbs.Add("pmplus_joblog");

            if (selectedDbs.Count == 0)
            {
                MessageBox.Show("DB 축소 작업을 진행할 데이터베이스를 최소 1개 이상 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(
                "선택한 데이터베이스에 대해 DB 축소(SHRINKDATABASE) 작업을 진행하시겠습니까?\n\n" +
                "- 대상 DB: " + string.Join(", ", selectedDbs.ToArray()) + "\n\n" +
                "※ 이 작업은 데이터베이스 파일의 빈 공간을 반환하여 디스크 여유 공간을 확보합니다.",
                "데이터베이스 축소 진행 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dr != DialogResult.Yes) return;

            SetShrinkButtonsEnabled(false);
            _txtDbRecoveryLog.Text = "";
            AppendRecoveryLog("============================================================\r\n");
            AppendRecoveryLog(string.Format("▶ [{0}] 데이터베이스 축소(SHRINKDATABASE) 작업 시작\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendRecoveryLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                RunDbShrinkDemo(selectedDbs);
            }
            else
            {
                RunDbShrinkProduction(selectedDbs);
            }
        }

        private void BtnShrinkLog_Click(object sender, EventArgs e)
        {
            List<string> selectedDbs = new List<string>();
            if (_chkDbPmMain.Checked) selectedDbs.Add("pm_main");
            if (_chkDbPmplusDums.Checked) selectedDbs.Add("pmplus_dums");
            if (_chkDbPmplusImage.Checked) selectedDbs.Add("pmplus_image");
            if (_chkDbPmplusJoblog.Checked) selectedDbs.Add("pmplus_joblog");

            if (selectedDbs.Count == 0)
            {
                MessageBox.Show("LOG 축소 작업을 진행할 데이터베이스를 최소 1개 이상 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(
                "선택한 데이터베이스에 대해 트랜잭션 LOG 축소(SHRINKFILE) 작업을 진행하시겠습니까?\n\n" +
                "- 대상 DB: " + string.Join(", ", selectedDbs.ToArray()) + "\n\n" +
                "※ 이 작업은 비정상적으로 비대해진 로그 파일(.ldf)을 축소하여 디스크 공간을 확보합니다.",
                "로그 파일 축소 진행 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dr != DialogResult.Yes) return;

            SetShrinkButtonsEnabled(false);
            _txtDbRecoveryLog.Text = "";
            AppendRecoveryLog("============================================================\r\n");
            AppendRecoveryLog(string.Format("▶ [{0}] 트랜잭션 로그 축소(SHRINKFILE) 작업 시작\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendRecoveryLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                RunLogShrinkDemo(selectedDbs);
            }
            else
            {
                RunLogShrinkProduction(selectedDbs);
            }
        }

        private void SetShrinkButtonsEnabled(bool enabled)
        {
            _btnRunDbRecovery.Enabled = enabled;
            _btnShrinkDb.Enabled = enabled;
            _btnShrinkLog.Enabled = enabled;
            if (_btnDropDrugUpdateDb != null) _btnDropDrugUpdateDb.Enabled = enabled;
            if (_btnDropDurUpdateDb != null) _btnDropDurUpdateDb.Enabled = enabled;
        }

        private void BtnDropDrugUpdateDb_Click(object sender, EventArgs e)
        {
            DropUpdateDatabase("PM_DRUGUPDATE");
        }

        private void BtnDropDurUpdateDb_Click(object sender, EventArgs e)
        {
            DropUpdateDatabase("PM_DURUPDATE");
        }

        private void DropUpdateDatabase(string dbName)
        {
            DialogResult dr = MessageBox.Show(
                string.Format("[{0}] 데이터베이스의 기존 연결을 강제 종료(SINGLE_USER)한 후 완전히 삭제(DROP DATABASE)하시겠습니까?\n\n※ 주의: 이 작업은 해당 임시 업데이트 DB를 SQL Server에서 영구 삭제합니다.", dbName),
                dbName + " 삭제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (dr != DialogResult.Yes) return;

            SetShrinkButtonsEnabled(false);
            _txtDbRecoveryLog.Text = "";
            AppendRecoveryLog("============================================================\r\n");
            AppendRecoveryLog(string.Format("▶ [{0}] [{1}] 데이터베이스 강제 연결 종료 및 삭제 시작\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), dbName));
            AppendRecoveryLog("============================================================\r\n");

            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                AppendRecoveryLog(string.Format("[가상 데모] USE master;\r\n"));
                AppendRecoveryLog(string.Format("[가상 데모] IF EXISTS (SELECT name FROM sys.databases WHERE name = '{0}')\r\n", dbName));
                AppendRecoveryLog(string.Format("[가상 데모] BEGIN\r\n"));
                AppendRecoveryLog(string.Format("[가상 데모]     ALTER DATABASE {0} SET SINGLE_USER WITH ROLLBACK IMMEDIATE;\r\n", dbName));
                AppendRecoveryLog(string.Format("[가상 데모]     DROP DATABASE {0};\r\n", dbName));
                AppendRecoveryLog(string.Format("[가상 데모] END\r\n"));
                AppendRecoveryLog(string.Format("✅ [가상 데모] {0} 데이터베이스가 성공적으로 삭제되었습니다.\r\n", dbName));
                SetShrinkButtonsEnabled(true);
                ShowToast(dbName + " 삭제 완료 (데모)", ColorEmerald);
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    string connStr = BuildConnectionString(false);
                    SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
                    builder.InitialCatalog = "master";

                    using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
                    {
                        conn.Open();

                        string checkSql = "SELECT COUNT(*) FROM sys.databases WHERE name = @dbName";
                        bool exists = false;
                        using (SqlCommand checkCmd = new SqlCommand(checkSql, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@dbName", dbName);
                            int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                            exists = (count > 0);
                        }

                        if (!exists)
                        {
                            this.BeginInvoke((Action)(() =>
                            {
                                AppendRecoveryLog(string.Format("ℹ️ [{0}] 데이터베이스가 SQL Server에 존재하지 않습니다. (이미 삭제됨)\r\n", dbName));
                                SetShrinkButtonsEnabled(true);
                                ShowToast(dbName + " 미존재 (이미 삭제됨)", ColorTextSec);
                                this.Cursor = Cursors.Default;
                            }));
                            return;
                        }

                        this.BeginInvoke((Action)(() =>
                        {
                            AppendRecoveryLog(string.Format("-> [{0}] 데이터베이스 존재 확인됨. 기존 연결 강제 종료(SINGLE_USER) 및 삭제 진행 중...\r\n", dbName));
                        }));

                        string dropSql = string.Format(@"
                            ALTER DATABASE {0} SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                            DROP DATABASE {0};", QuoteSqlName(dbName));

                        using (SqlCommand dropCmd = new SqlCommand(dropSql, conn))
                        {
                            dropCmd.CommandTimeout = 120;
                            dropCmd.ExecuteNonQuery();
                        }

                        this.BeginInvoke((Action)(() =>
                        {
                            AppendRecoveryLog(string.Format("✅ [{0}] 데이터베이스 연결이 강제 종료되고 성공적으로 삭제(DROP DATABASE)되었습니다.\r\n", dbName));
                            AppendRecoveryLog(string.Format("   - 작업 완료 일시: {0}\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                            SetShrinkButtonsEnabled(true);
                            ShowToast(dbName + " 삭제 성공", ColorEmerald);
                            MessageBox.Show(string.Format("[{0}] 데이터베이스가 성공적으로 삭제되었습니다.", dbName), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.Cursor = Cursors.Default;
                        }));
                    }
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        AppendRecoveryLog(string.Format("❌ [{0}] 삭제 실패:\r\n{1}\r\n", dbName, ex.Message));
                        SetShrinkButtonsEnabled(true);
                        ShowToast(dbName + " 삭제 실패", ColorAlarm);
                        MessageBox.Show(string.Format("[{0}] 데이터베이스 삭제 중 오류 발생:\n\n{1}", dbName, ex.Message), "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Cursor = Cursors.Default;
                    }));
                }
            });
        }

        private void AppendRecoveryLog(string msg)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((Action)(() => AppendRecoveryLog(msg)));
            }
            else
            {
                _txtDbRecoveryLog.AppendText(msg);
                _txtDbRecoveryLog.SelectionStart = _txtDbRecoveryLog.Text.Length;
                _txtDbRecoveryLog.ScrollToCaret();
            }
        }

        private void RunDbRecoveryDemo(List<string> dbs)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    foreach (string db in dbs)
                    {
                        AppendRecoveryLog(string.Format("\r\n[DB: {0}] 응급 처리 개시...\r\n", db));
                        System.Threading.Thread.Sleep(800);

                        AppendRecoveryLog(string.Format("  - 1단계: EXEC sp_resetstatus {0};\r\n", db));
                        System.Threading.Thread.Sleep(500);

                        AppendRecoveryLog(string.Format("  - 2단계: ALTER DATABASE {0} SET EMERGENCY;\r\n", db));
                        System.Threading.Thread.Sleep(500);

                        AppendRecoveryLog(string.Format("  - 3단계: DBCC checkdb('{0}') 실행 중...\r\n", db));
                        System.Threading.Thread.Sleep(1200);
                        AppendRecoveryLog(string.Format("    ➔ DBCC checkdb 결과: {0} 내 0개의 할당 오류 및 0개의 일관성 오류 발견.\r\n", db));

                        AppendRecoveryLog(string.Format("  - 4단계: ALTER DATABASE {0} SET SINGLE_USER WITH ROLLBACK IMMEDIATE;\r\n", db));
                        System.Threading.Thread.Sleep(600);

                        AppendRecoveryLog(string.Format("  - 5단계: DBCC CheckDB ('{0}', REPAIR_ALLOW_DATA_LOSS) 실행 중...\r\n", db));
                        System.Threading.Thread.Sleep(1500);
                        AppendRecoveryLog(string.Format("    ➔ DBCC 복구 완료: 복구 엔진이 {0}의 테이블 손상 부위를 처리했습니다. (가상 손실: 0건)\r\n", db));

                        AppendRecoveryLog(string.Format("  - 6단계: ALTER DATABASE {0} SET MULTI_USER;\r\n", db));
                        System.Threading.Thread.Sleep(500);

                        AppendRecoveryLog(string.Format("▶ [DB: {0}] 응급 복구 성공 완료!\r\n", db));
                    }

                    AppendRecoveryLog("\r\n============================================================\r\n");
                    AppendRecoveryLog(string.Format("▶ [{0}] [데모] 모든 선택된 DB 응급 시퀀스가 성공적으로 마쳐졌습니다.\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                    AppendRecoveryLog("============================================================\r\n");
                }
                catch (Exception ex)
                {
                    AppendRecoveryLog(string.Format("\r\n[에러 발생] {0}\r\n", ex.Message));
                }
                finally
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        _btnRunDbRecovery.Enabled = true;
                        _btnRunDbRecovery.Text = "⚡ DB 응급 복구 실행";
                        MessageBox.Show("DB 응급 복구 데모가 완료되었습니다. 로그를 확인해 주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
            });
        }

        private void RunDbRecoveryProduction(List<string> dbs)
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(BuildConnectionString(false));
            builder.InitialCatalog = "master";
            string masterConnStr = builder.ConnectionString;

            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(masterConnStr))
                    {
                        conn.Open();

                        foreach (string db in dbs)
                        {
                            AppendRecoveryLog(string.Format("\r\n[DB: {0}] SQL Server 응급 복구 작업을 개시합니다...\r\n", db));

                            try
                            {
                                AppendRecoveryLog(string.Format("  - 1단계: EXEC sp_resetstatus '{0}' 실행...\r\n", db));
                                string sql = string.Format("EXEC sp_resetstatus [{0}]", db.Replace("]", "]]"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 0;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog("    ➔ 완료\r\n");
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("    ➔ [경고] 1단계 건너뜀/실패: {0}\r\n", ex.Message));
                            }

                            try
                            {
                                AppendRecoveryLog(string.Format("  - 2단계: ALTER DATABASE [{0}] SET EMERGENCY 실행...\r\n", db));
                                string sql = string.Format("ALTER DATABASE [{0}] SET EMERGENCY", db.Replace("]", "]]"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 0;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog("    ➔ 완료\r\n");
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("    ➔ [에러] 2단계 실패: {0}\r\n", ex.Message));
                                throw;
                            }

                            try
                            {
                                AppendRecoveryLog(string.Format("  - 3단계: DBCC CHECKDB ('{0}') 무결성 정밀 진단 중 (대기)...\r\n", db));
                                string sql = string.Format("DBCC CHECKDB ('{0}') WITH NO_INFOMSGS, ALL_ERRORMSGS", db.Replace("'", "''"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 0;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog("    ➔ 진단 완료 (무결성 검사 성공)\r\n");
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("    ➔ [경고] 3단계 검사 중 이관 필요 오류 발견: {0}\r\n", ex.Message));
                            }

                            try
                            {
                                AppendRecoveryLog(string.Format("  - 4단계: ALTER DATABASE [{0}] SET SINGLE_USER 모드 전환...\r\n", db));
                                string sql = string.Format("ALTER DATABASE [{0}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE", db.Replace("]", "]]"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 0;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog("    ➔ 단일 사용자 모드 활성화 완료\r\n");
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("    ➔ [에러] 4단계 진입 실패 (타 세션 물림): {0}\r\n", ex.Message));
                                throw;
                            }

                            try
                            {
                                AppendRecoveryLog(string.Format("  - 5단계: DBCC CHECKDB ('{0}', REPAIR_ALLOW_DATA_LOSS) 복구 강제 실행 (대기)...\r\n", db));
                                string sql = string.Format("DBCC CHECKDB ('{0}', REPAIR_ALLOW_DATA_LOSS)", db.Replace("'", "''"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 0;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog("    ➔ 깨진 데이터 복구 및 무결성 패치 완료\r\n");
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("    ➔ [경고] 5단계 복구 완료되었으나 일부 오류 잔존할 수 있음: {0}\r\n", ex.Message));
                            }
                            finally
                            {
                                try
                                {
                                    AppendRecoveryLog(string.Format("  - 6단계: ALTER DATABASE [{0}] SET MULTI_USER 다중 사용자 모드 원복...\r\n", db));
                                    string sql = string.Format("ALTER DATABASE [{0}] SET MULTI_USER", db.Replace("]", "]]"));
                                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                                    {
                                        cmd.CommandTimeout = 0;
                                        cmd.ExecuteNonQuery();
                                    }
                                    AppendRecoveryLog("    ➔ 다중 사용자 모드 원복 완료\r\n");
                                }
                                catch (Exception ex)
                                {
                                    AppendRecoveryLog(string.Format("    ➔ [경고] 6단계 복귀 실패 (수동 제어 필요): {0}\r\n", ex.Message));
                                }
                            }

                            AppendRecoveryLog(string.Format("▶ [DB: {0}] 데이터베이스 복구 처리 완료!\r\n", db));
                        }
                    }

                    AppendRecoveryLog("\r\n============================================================\r\n");
                    AppendRecoveryLog(string.Format("▶ [{0}] 모든 실제 DB 응급 복구 절차가 정상적으로 마쳐졌습니다.\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                    AppendRecoveryLog("============================================================\r\n");
                }
                catch (Exception ex)
                {
                    AppendRecoveryLog(string.Format("\r\n[응급 중단 에러 발생] {0}\r\n", ex.Message));
                }
                finally
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        _btnRunDbRecovery.Enabled = true;
                        _btnRunDbRecovery.Text = "⚡ DB 응급 복구 실행";
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("DB 응급 복구 및 무결성 검사가 마쳐졌습니다.\r\n하단의 상세 콘솔 로그를 확인해주십시오.", "작업 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
            });
        }

        private void RunDbShrinkDemo(List<string> dbs)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                foreach (string db in dbs)
                {
                    System.Threading.Thread.Sleep(800);
                    AppendRecoveryLog(string.Format("✔ [DB: {0}] DBCC SHRINKDATABASE 가상 완료!\r\n", db));
                }
                
                this.BeginInvoke((Action)(() =>
                {
                    SetShrinkButtonsEnabled(true);
                    AppendRecoveryLog("\r\n🎉 데이터베이스 축소 완료 (데모 모드)\r\n");
                    MessageBox.Show("DB 축소(데모)가 완료되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
        }

        private void RunLogShrinkDemo(List<string> dbs)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                foreach (string db in dbs)
                {
                    System.Threading.Thread.Sleep(800);
                    AppendRecoveryLog(string.Format("✔ [DB: {0}] DBCC SHRINKFILE (2, 1) 가상 완료!\r\n", db));
                }
                
                this.BeginInvoke((Action)(() =>
                {
                    SetShrinkButtonsEnabled(true);
                    AppendRecoveryLog("\r\n🎉 로그 축소 완료 (데모 모드)\r\n");
                    MessageBox.Show("LOG 축소(데모)가 완료되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
        }

        private void RunDbShrinkProduction(List<string> dbs)
        {
            string connStr = BuildConnectionString(false);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
            builder.InitialCatalog = "master";
            string masterConnStr = builder.ConnectionString;

            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(masterConnStr))
                    {
                        conn.Open();
                        foreach (string db in dbs)
                        {
                            AppendRecoveryLog(string.Format("\r\n[DB: {0}] 데이터베이스 축소 실행 중...\r\n", db));
                            try
                            {
                                string sql = string.Format("DBCC SHRINKDATABASE ([{0}])", db.Replace("]", "]]"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 600; // 10 minutes timeout
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog(string.Format("  ➔ ✔ [DB: {0}] 축소 성공!\r\n", db));
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("  ➔ ❌ [DB: {0}] 축소 실패: {1}\r\n", db, ex.Message));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendRecoveryLog(string.Format("\r\n❌ 오류 발생: {0}\r\n", ex.Message));
                }
                finally
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        SetShrinkButtonsEnabled(true);
                        this.Cursor = Cursors.Default;
                        AppendRecoveryLog(string.Format("\r\n✔ [{0}] 데이터베이스 축소 작업 종료.\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                        MessageBox.Show("선택한 데이터베이스 축소 작업이 종료되었습니다.\r\n하단 로그 콘솔에서 결과를 확인하십시오.", "작업 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
            });
        }

        private void RunLogShrinkProduction(List<string> dbs)
        {
            string connStr = BuildConnectionString(false);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
            builder.InitialCatalog = "master";
            string masterConnStr = builder.ConnectionString;

            this.Cursor = Cursors.WaitCursor;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(masterConnStr))
                    {
                        conn.Open();
                        foreach (string db in dbs)
                        {
                            AppendRecoveryLog(string.Format("\r\n[DB: {0}] 트랜잭션 로그 축소 실행 중...\r\n", db));
                            try
                            {
                                string sql = string.Format("USE [{0}]; DBCC SHRINKFILE (2, 1);", db.Replace("]", "]]"));
                                using (SqlCommand cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.CommandTimeout = 600;
                                    cmd.ExecuteNonQuery();
                                }
                                AppendRecoveryLog(string.Format("  ➔ ✔ [DB: {0}] 로그 축소 성공!\r\n", db));
                            }
                            catch (Exception ex)
                            {
                                AppendRecoveryLog(string.Format("  ➔ ❌ [DB: {0}] 로그 축소 실패: {1}\r\n", db, ex.Message));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendRecoveryLog(string.Format("\r\n❌ 오류 발생: {0}\r\n", ex.Message));
                }
                finally
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        SetShrinkButtonsEnabled(true);
                        this.Cursor = Cursors.Default;
                        AppendRecoveryLog(string.Format("\r\n✔ [{0}] 트랜잭션 로그 축소 작업 종료.\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                        MessageBox.Show("선택한 트랜잭션 로그 축소 작업이 종료되었습니다.\r\n하단 로그 콘솔에서 결과를 확인하십시오.", "작업 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
            });
        }

        // ==========================================
        // CRUD Data Management & Hashing Implementation
        // ==========================================

        // SHA-512 Hash Generator (128 char upper-case hex string)
        private string GetSHA512Hash(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            using (System.Security.Cryptography.SHA512 sha = System.Security.Cryptography.SHA512.Create())
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = sha.ComputeHash(bytes);
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }
                return sb.ToString();
            }
        }

        // Initialize Data Management (5th) Tab Layout
        private void InitializeDataManagementTab()
        {
            // 부가 데이터 관리 ➡️ 기초 데이터 관리로 이름 변경
            _tabDataManagement = new TabPage
            {
                Text = "💾 기초 데이터 관리",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabDataManagement);

            // 내부 서브 TabControl
            _subTabDataManagement = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            _tabDataManagement.Controls.Add(_subTabDataManagement);
            _subTabDataManagement.SelectedIndexChanged += (s, e) =>
            {
                BeginInvoke((Action)(() =>
                {
                    if (_splitUser  != null && _splitUser.Visible)  NormalizeDataManagementSplit(_splitUser, ref _distUser);
                    if (_splitCard  != null && _splitCard.Visible)  NormalizeDataManagementSplit(_splitCard, ref _distCard);
                    if (_splitLabel != null && _splitLabel.Visible) NormalizeDataManagementSplit(_splitLabel, ref _distLabel);
                    if (_splitRx    != null && _splitRx.Visible)    NormalizeRightPanelSplit(_splitRx, ref _distRx, 340, 360);
                }));
            };

            // ------------------ 서브 탭 1: 사용자 ID 관리 ------------------
            TabPage tabUser = new TabPage
            {
                Text = "👤 사용자 ID 관리 (TBSIM000_09)",
                BackColor = ColorBgMain
            };
            _subTabDataManagement.TabPages.Add(tabUser);
            
            // Top 검색 패널
            Panel pnlUserSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            tabUser.Controls.Add(pnlUserSearch);

            Label lblUserSearchId = new Label { Text = "사용자 ID", Location = new Point(15, 20), Size = new Size(70, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtUserSearchId = new TextBox { Location = new Point(90, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            Label lblUserSearchName = new Label { Text = "이름", Location = new Point(230, 20), Size = new Size(40, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtUserSearchName = new TextBox { Location = new Point(280, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            
            _btnUserSearch = new Button
            {
                Text = "🔍 사용자 조회",
                Location = new Point(420, 14),
                Size = new Size(120, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnUserSearch.Click += BtnUserSearch_Click;
            pnlUserSearch.Controls.Add(lblUserSearchId);
            pnlUserSearch.Controls.Add(_txtUserSearchId);
            pnlUserSearch.Controls.Add(lblUserSearchName);
            pnlUserSearch.Controls.Add(_txtUserSearchName);
            pnlUserSearch.Controls.Add(_btnUserSearch);
            _splitUser = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = _distUser,
                BackColor = ColorBorder
            };
            tabUser.Controls.Add(_splitUser);
            _splitUser.BringToFront();
            _splitUser.SplitterMoved += (s, e) => { _distUser = _splitUser.SplitterDistance; SaveConfig(); };
            _splitUser.Resize += (s, e) => NormalizeDataManagementSplit(_splitUser, ref _distUser);
            NormalizeDataManagementSplit(_splitUser, ref _distUser);

            _dgvUsers = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvUsers.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvUsers.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvUsers.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvUsers.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvUsers.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvUsers.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvUsers.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvUsers.CellClick += DgvUsers_CellClick;
            _splitUser.Panel1.Controls.Add(_dgvUsers); // Panel1(왼쪽)에 목록 배치

            Panel pnlUserForm = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(20)
            };
            _splitUser.Panel2.Controls.Add(pnlUserForm); // Panel2(오른쪽)에 입력 폼 배치

            Label lblFormTitle1 = new Label { Text = "👤 사용자 정보 관리", Location = new Point(20, 15), Size = new Size(200, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold), ForeColor = ColorIndigo };
            pnlUserForm.Controls.Add(lblFormTitle1);

            int uy = 55;
            Label lblUserId = new Label { Text = "사용자 ID *", Location = new Point(20, uy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtUserId = new TextBox { Location = new Point(110, uy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlUserForm.Controls.Add(lblUserId); pnlUserForm.Controls.Add(_txtUserId);

            uy += 35;
            Label lblUserNm = new Label { Text = "사용자 이름 *", Location = new Point(20, uy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtUserNm = new TextBox { Location = new Point(110, uy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlUserForm.Controls.Add(lblUserNm); pnlUserForm.Controls.Add(_txtUserNm);

            uy += 35;
            Label lblUserPwd = new Label { Text = "비밀번호", Location = new Point(20, uy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtUserPwd = new TextBox { Location = new Point(110, uy - 3), Size = new Size(180, 25), PasswordChar = '●', BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlUserForm.Controls.Add(lblUserPwd); pnlUserForm.Controls.Add(_txtUserPwd);

            uy += 35;
            Label lblUserDeptCd = new Label { Text = "부서 코드", Location = new Point(20, uy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtUserDeptCd = new TextBox { Location = new Point(110, uy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlUserForm.Controls.Add(lblUserDeptCd); pnlUserForm.Controls.Add(_txtUserDeptCd);

            uy += 35;
            Label lblUserLicNo = new Label { Text = "약사면허번호", Location = new Point(20, uy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtUserLicNo = new TextBox { Location = new Point(110, uy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlUserForm.Controls.Add(lblUserLicNo); pnlUserForm.Controls.Add(_txtUserLicNo);

            uy += 45;
            _btnUserAdd = new Button { Text = "➕ 추가", Location = new Point(20, uy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorEmerald, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnUserUpdate = new Button { Text = "✏️ 수정", Location = new Point(95, uy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorIndigo, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnUserDelete = new Button { Text = "🗑️ 삭제", Location = new Point(170, uy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorAlarm, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnUserClear = new Button { Text = "🔄 비우기", Location = new Point(245, uy), Size = new Size(80, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorBorder, ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };

            _btnUserAdd.FlatAppearance.BorderSize = 0; _btnUserUpdate.FlatAppearance.BorderSize = 0; _btnUserDelete.FlatAppearance.BorderSize = 0; _btnUserClear.FlatAppearance.BorderSize = 0;

            _btnUserAdd.Click += BtnUserAdd_Click;
            _btnUserUpdate.Click += BtnUserUpdate_Click;
            _btnUserDelete.Click += BtnUserDelete_Click;
            _btnUserClear.Click += (s, e) => ClearUserForm();

            pnlUserForm.Controls.Add(_btnUserAdd);
            pnlUserForm.Controls.Add(_btnUserUpdate);
            pnlUserForm.Controls.Add(_btnUserDelete);
            pnlUserForm.Controls.Add(_btnUserClear);

            // ------------------ 서브 탭 2: 카드결제내역 관리 ------------------
            TabPage tabCard = new TabPage
            {
                Text = "💳 카드결제내역 관리 (tbsir000_01)",
                BackColor = ColorBgMain
            };
            _subTabDataManagement.TabPages.Add(tabCard);
            
            // Top 검색 패널
            Panel pnlCardSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            tabCard.Controls.Add(pnlCardSearch);

            Label lblCardSearchChart = new Label { Text = "차트번호", Location = new Point(15, 20), Size = new Size(60, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtCardSearchChart = new TextBox { Location = new Point(80, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            Label lblCardSearchDate = new Label { Text = "수납일자", Location = new Point(220, 20), Size = new Size(60, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtCardSearchDate = new TextBox { Location = new Point(290, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            
            _btnCardSearch = new Button
            {
                Text = "🔍 내역 조회",
                Location = new Point(430, 14),
                Size = new Size(120, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnCardSearch.FlatAppearance.BorderSize = 0;
            _btnCardSearch.Click += BtnCardSearch_Click;

            pnlCardSearch.Controls.Add(lblCardSearchChart);
            pnlCardSearch.Controls.Add(_txtCardSearchChart);
            pnlCardSearch.Controls.Add(lblCardSearchDate);
            pnlCardSearch.Controls.Add(_txtCardSearchDate);
            pnlCardSearch.Controls.Add(_btnCardSearch);

            // SplitContainer
            _splitCard = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = _distCard,
                BackColor = ColorBorder
            };
            tabCard.Controls.Add(_splitCard);
            _splitCard.BringToFront();
            _splitCard.SplitterMoved += (s, e) => { _distCard = _splitCard.SplitterDistance; SaveConfig(); };
            _splitCard.Resize += (s, e) => NormalizeDataManagementSplit(_splitCard, ref _distCard);
            NormalizeDataManagementSplit(_splitCard, ref _distCard);

            _dgvCardPays = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvCardPays.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvCardPays.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvCardPays.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvCardPays.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvCardPays.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvCardPays.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvCardPays.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvCardPays.CellClick += DgvCardPays_CellClick;
            _splitCard.Panel1.Controls.Add(_dgvCardPays); // Panel1(왼쪽)에 목록 배치

            Panel pnlCardForm = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(20)
            };
            _splitCard.Panel2.Controls.Add(pnlCardForm); // Panel2(오른쪽)에 입력 폼 배치

            Label lblFormTitle2 = new Label { Text = "💳 카드결제 정보 관리", Location = new Point(20, 15), Size = new Size(200, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold), ForeColor = ColorIndigo };
            pnlCardForm.Controls.Add(lblFormTitle2);

            int cy = 55;
            Label lblCardSlipSeq = new Label { Text = "일련번호 *", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardSlipSeq = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle, ReadOnly = true };
            pnlCardForm.Controls.Add(lblCardSlipSeq); pnlCardForm.Controls.Add(_txtCardSlipSeq);

            cy += 35;
            Label lblCardRecpDt = new Label { Text = "수납일자 *", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardRecpDt = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardRecpDt); pnlCardForm.Controls.Add(_txtCardRecpDt);

            cy += 35;
            Label lblCardChrtNo = new Label { Text = "차트번호 *", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardChrtNo = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardChrtNo); pnlCardForm.Controls.Add(_txtCardChrtNo);

            cy += 35;
            Label lblCardCoNm = new Label { Text = "카드사명", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardCoNm = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardCoNm); pnlCardForm.Controls.Add(_txtCardCoNm);

            cy += 35;
            Label lblCardAmt = new Label { Text = "카드금액 *", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardAmt = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardAmt); pnlCardForm.Controls.Add(_txtCardAmt);

            cy += 35;
            Label lblCardAdmNo = new Label { Text = "승인번호", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardAdmNo = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardAdmNo); pnlCardForm.Controls.Add(_txtCardAdmNo);

            cy += 35;
            Label lblCardNo = new Label { Text = "카드번호", Location = new Point(20, cy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtCardNo = new TextBox { Location = new Point(110, cy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlCardForm.Controls.Add(lblCardNo); pnlCardForm.Controls.Add(_txtCardNo);

            cy += 45;
            _btnCardAdd = new Button { Text = "➕ 추가", Location = new Point(20, cy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorEmerald, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnCardUpdate = new Button { Text = "✏️ 수정", Location = new Point(95, cy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorIndigo, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnCardDelete = new Button { Text = "🗑️ 삭제", Location = new Point(170, cy), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorAlarm, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnCardClear = new Button { Text = "🔄 비우기", Location = new Point(245, cy), Size = new Size(80, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorBorder, ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnCardAdd.FlatAppearance.BorderSize = 0; _btnCardUpdate.FlatAppearance.BorderSize = 0; _btnCardDelete.FlatAppearance.BorderSize = 0; _btnCardClear.FlatAppearance.BorderSize = 0;

            _btnCardAdd.Click += BtnCardAdd_Click;
            _btnCardUpdate.Click += BtnCardUpdate_Click;
            _btnCardDelete.Click += BtnCardDelete_Click;
            _btnCardClear.Click += (s, e) => ClearCardForm();

            pnlCardForm.Controls.Add(_btnCardAdd);
            pnlCardForm.Controls.Add(_btnCardUpdate);
            pnlCardForm.Controls.Add(_btnCardDelete);
            pnlCardForm.Controls.Add(_btnCardClear);

            // ------------------ 서브 탭 3: 라벨출력정보 관리 ------------------
            TabPage tabLabel = new TabPage
            {
                Text = "🏷️ 라벨출력정보 관리 (TBSIM040_43)",
                BackColor = ColorBgMain
            };
            _subTabDataManagement.TabPages.Add(tabLabel);
            
            // Top 검색 패널
            Panel pnlLabelSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            tabLabel.Controls.Add(pnlLabelSearch);

            Label lblLabelSearchCode = new Label { Text = "약품코드", Location = new Point(15, 20), Size = new Size(60, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtLabelSearchCode = new TextBox { Location = new Point(80, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            Label lblLabelSearchName = new Label { Text = "약품명", Location = new Point(220, 20), Size = new Size(50, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtLabelSearchName = new TextBox { Location = new Point(285, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            
            _btnLabelSearch = new Button
            {
                Text = "🔍 내역 조회",
                Location = new Point(425, 14),
                Size = new Size(120, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnLabelSearch.FlatAppearance.BorderSize = 0;
            _btnLabelSearch.Click += BtnLabelSearch_Click;

            pnlLabelSearch.Controls.Add(lblLabelSearchCode);
            pnlLabelSearch.Controls.Add(_txtLabelSearchCode);
            pnlLabelSearch.Controls.Add(lblLabelSearchName);
            pnlLabelSearch.Controls.Add(_txtLabelSearchName);
            pnlLabelSearch.Controls.Add(_btnLabelSearch);

            // SplitContainer
            _splitLabel = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = _distLabel,
                BackColor = ColorBorder
            };
            tabLabel.Controls.Add(_splitLabel);
            _splitLabel.BringToFront();
            _splitLabel.SplitterMoved += (s, e) => { _distLabel = _splitLabel.SplitterDistance; SaveConfig(); };
            _splitLabel.Resize += (s, e) => NormalizeDataManagementSplit(_splitLabel, ref _distLabel);
            NormalizeDataManagementSplit(_splitLabel, ref _distLabel);

            _dgvLabelInfos = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvLabelInfos.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvLabelInfos.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvLabelInfos.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvLabelInfos.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvLabelInfos.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvLabelInfos.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvLabelInfos.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvLabelInfos.CellClick += DgvLabelInfos_CellClick;
            _splitLabel.Panel1.Controls.Add(_dgvLabelInfos); // Panel1(왼쪽)에 목록 배치

            Panel pnlLabelForm = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(20),
                AutoScroll = true
            };
            _splitLabel.Panel2.Controls.Add(pnlLabelForm); // Panel2(오른쪽)에 입력 폼 배치

            Label lblFormTitle3 = new Label { Text = "🏷️ 라벨출력 정보 관리", Location = new Point(20, 15), Size = new Size(200, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold), ForeColor = ColorIndigo };
            pnlLabelForm.Controls.Add(lblFormTitle3);

            int ly = 55;
            Label lblLabelDrugCode = new Label { Text = "약품코드 *", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelDrugCode = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelDrugCode); pnlLabelForm.Controls.Add(_txtLabelDrugCode);

            ly += 35;
            Label lblLabelDrug = new Label { Text = "약품명", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelDrug = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelDrug); pnlLabelForm.Controls.Add(_txtLabelDrug);

            ly += 35;
            Label lblLabelDan = new Label { Text = "단위", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelDan = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelDan); pnlLabelForm.Controls.Add(_txtLabelDan);

            ly += 35;
            Label lblLabelSave = new Label { Text = "보관방법", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelSave = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelSave); pnlLabelForm.Controls.Add(_txtLabelSave);

            ly += 35;
            Label lblLabelPrintOp = new Label { Text = "출력옵션", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelPrintOp = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelPrintOp); pnlLabelForm.Controls.Add(_txtLabelPrintOp);

            ly += 35;
            Label lblLabelInputOp = new Label { Text = "입력옵션", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelInputOp = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelInputOp); pnlLabelForm.Controls.Add(_txtLabelInputOp);

            ly += 35;
            Label lblLabelEffct = new Label { Text = "효능효과", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelEffct = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelEffct); pnlLabelForm.Controls.Add(_txtLabelEffct);

            ly += 35;
            Label lblLabelComment = new Label { Text = "코멘트", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelComment = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelComment); pnlLabelForm.Controls.Add(_txtLabelComment);

            ly += 35;
            Label lblLabelSampleUp = new Label { Text = "샘플구분 *", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelSampleUp = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelSampleUp); pnlLabelForm.Controls.Add(_txtLabelSampleUp);

            ly += 35;
            Label lblLabelEffctUnit = new Label { Text = "효능단위", Location = new Point(20, ly), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtLabelEffctUnit = new TextBox { Location = new Point(110, ly - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlLabelForm.Controls.Add(lblLabelEffctUnit); pnlLabelForm.Controls.Add(_txtLabelEffctUnit);

            ly += 45;
            _btnLabelAdd = new Button { Text = "➕ 추가", Location = new Point(20, ly), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorEmerald, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnLabelUpdate = new Button { Text = "✏️ 수정", Location = new Point(95, ly), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorIndigo, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnLabelDelete = new Button { Text = "🗑️ 삭제", Location = new Point(170, ly), Size = new Size(70, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorAlarm, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnLabelClear = new Button { Text = "🔄 비우기", Location = new Point(245, ly), Size = new Size(80, 30), FlatStyle = FlatStyle.Flat, BackColor = ColorBorder, ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnLabelAdd.FlatAppearance.BorderSize = 0; _btnLabelUpdate.FlatAppearance.BorderSize = 0; _btnLabelDelete.FlatAppearance.BorderSize = 0; _btnLabelClear.FlatAppearance.BorderSize = 0;

            _btnLabelAdd.Click += BtnLabelAdd_Click;
            _btnLabelUpdate.Click += BtnLabelUpdate_Click;
            _btnLabelDelete.Click += BtnLabelDelete_Click;
            _btnLabelClear.Click += (s, e) => ClearLabelForm();

            pnlLabelForm.Controls.Add(_btnLabelAdd);
            pnlLabelForm.Controls.Add(_btnLabelUpdate);
            pnlLabelForm.Controls.Add(_btnLabelDelete);
            pnlLabelForm.Controls.Add(_btnLabelClear);

            // ------------------ 상위 탭: 재고관련 ------------------
            _tabInventoryManagement = new TabPage
            {
                Text = "📦 재고관련",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabInventoryManagement);

            _subTabInventoryManagement = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            _tabInventoryManagement.Controls.Add(_subTabInventoryManagement);

            TabPage tabInventoryCleanup = new TabPage
            {
                Text = "🧹 약품 정보 정리",
                BackColor = ColorBgMain
            };
            _subTabInventoryManagement.TabPages.Add(tabInventoryCleanup);

            TabPage tabStockMovementErrors = new TabPage
            {
                Text = "⚠️ 입출고 오류",
                BackColor = ColorBgMain
            };
            _subTabInventoryManagement.TabPages.Add(tabStockMovementErrors);

            // Top 검색 패널
            Panel pnlInvSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 96,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            tabInventoryCleanup.Controls.Add(pnlInvSearch);

            Label lblInvSearch = new Label { Text = "약품코드/명", Location = new Point(15, 20), Size = new Size(85, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtInventorySearch = new TextBox { Location = new Point(105, 17), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _chkInventoryNoNameOnly = new CheckBox { Text = "약품명이 없는 약품만 보기", Location = new Point(305, 17), Size = new Size(200, 25), ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _chkInventoryExcludeZeroStock = new CheckBox { Text = "재고합계가 0인 것은 제외", Location = new Point(305, 54), Size = new Size(210, 25), ForeColor = ColorTextMain, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };

            _btnInventorySearch = new Button
            {
                Text = "🔍 재고 조회",
                Location = new Point(520, 14),
                Size = new Size(120, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnInventorySearch.FlatAppearance.BorderSize = 0;
            _btnInventorySearch.Click += BtnInventorySearch_Click;

            _btnInvBatchDelete = new Button
            {
                Text = "🗑️ 이름 없는 재고 0 일괄 삭제",
                Location = new Point(650, 14),
                Size = new Size(210, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnInvBatchDelete.FlatAppearance.BorderSize = 0;
            _btnInvBatchDelete.Click += BtnInvBatchDelete_Click;

            _btnInvCleanDupBarcodes = new Button
            {
                Text = "🗑️ 중복 바코드 정리",
                Location = new Point(870, 14),
                Size = new Size(180, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnInvCleanDupBarcodes.FlatAppearance.BorderSize = 0;
            _btnInvCleanDupBarcodes.Click += BtnInvCleanDupBarcodes_Click;

            _btnDurakanAudit = new Button
            {
                Text = "🧪 듀락칸 500mL 오류 검사",
                Location = new Point(15, 52),
                Size = new Size(230, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnDurakanAudit.FlatAppearance.BorderSize = 0;
            _btnDurakanAudit.Click += BtnDurakanAudit_Click;

            pnlInvSearch.Controls.Add(lblInvSearch);
            pnlInvSearch.Controls.Add(_txtInventorySearch);
            pnlInvSearch.Controls.Add(_chkInventoryNoNameOnly);
            pnlInvSearch.Controls.Add(_chkInventoryExcludeZeroStock);
            pnlInvSearch.Controls.Add(_btnInventorySearch);
            pnlInvSearch.Controls.Add(_btnInvBatchDelete);
            pnlInvSearch.Controls.Add(_btnInvCleanDupBarcodes);

            // SplitContainer
            _splitInventory = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = 650,
                BackColor = ColorBorder
            };
            tabInventoryCleanup.Controls.Add(_splitInventory);
            _splitInventory.BringToFront();
            _splitInventory.Resize += (s, e) => NormalizeRightPanelSplit(_splitInventory, 320, 760);
            NormalizeRightPanelSplit(_splitInventory, 320, 760);

            // DataGridView
            _dgvInventory = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvInventory.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvInventory.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvInventory.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvInventory.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvInventory.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvInventory.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvInventory.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvInventory.CellClick += DgvInventory_CellClick;
            _splitInventory.Panel1.Controls.Add(_dgvInventory);

            // Right Panel (Form)
            Panel pnlInvForm = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(20),
                AutoScroll = true
            };
            _splitInventory.Panel2.Controls.Add(pnlInvForm);

            Label lblInvFormTitle = new Label { Text = "📝 약품 정보 수정", Location = new Point(20, 15), Size = new Size(200, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold), ForeColor = ColorIndigo };
            pnlInvForm.Controls.Add(lblInvFormTitle);

            int iy = 55;
            Label lblInvFormDrugCode = new Label { Text = "약품코드", Location = new Point(20, iy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtInvFormDrugCode = new TextBox { Location = new Point(110, iy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle, ReadOnly = true };
            pnlInvForm.Controls.Add(lblInvFormDrugCode); pnlInvForm.Controls.Add(_txtInvFormDrugCode);

            iy += 35;
            Label lblInvFormBarcode = new Label { Text = "바코드", Location = new Point(20, iy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtInvFormBarcode = new TextBox { Location = new Point(110, iy - 3), Size = new Size(115, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle, ReadOnly = true };
            _btnInvBarcodeSearchWeb = new Button
            {
                Text = "🌐 검색",
                Location = new Point(230, iy - 3),
                Size = new Size(60, 25),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _btnInvBarcodeSearchWeb.FlatAppearance.BorderSize = 0;
            _btnInvBarcodeSearchWeb.Click += BtnInvBarcodeSearchWeb_Click;
            pnlInvForm.Controls.Add(lblInvFormBarcode); 
            pnlInvForm.Controls.Add(_txtInvFormBarcode);
            pnlInvForm.Controls.Add(_btnInvBarcodeSearchWeb);

            iy += 35;
            Label lblInvFormDrugName = new Label { Text = "약품명 *", Location = new Point(20, iy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtInvFormDrugName = new TextBox { Location = new Point(110, iy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlInvForm.Controls.Add(lblInvFormDrugName); pnlInvForm.Controls.Add(_txtInvFormDrugName);

            iy += 35;
            Label lblInvFormManufacturer = new Label { Text = "제조회사", Location = new Point(20, iy), Size = new Size(80, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _txtInvFormManufacturer = new TextBox { Location = new Point(110, iy - 3), Size = new Size(180, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            pnlInvForm.Controls.Add(lblInvFormManufacturer); pnlInvForm.Controls.Add(_txtInvFormManufacturer);

            iy += 35;
            _lblInvFormSuggest = new Label { Text = "", Location = new Point(20, iy), Size = new Size(270, 45), ForeColor = ColorEmerald, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            pnlInvForm.Controls.Add(_lblInvFormSuggest);

            iy += 50;
            _btnInvFormUpdate = new Button
            {
                Text = "✏️ 정보 저장",
                Location = new Point(20, iy),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnInvFormUpdate.FlatAppearance.BorderSize = 0;
            _btnInvFormUpdate.Click += BtnInvFormUpdate_Click;
            pnlInvForm.Controls.Add(_btnInvFormUpdate);

            _btnInvFormDelete = new Button
            {
                Text = "🗑️ 약품 삭제",
                Location = new Point(140, iy),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnInvFormDelete.FlatAppearance.BorderSize = 0;
            _btnInvFormDelete.Click += BtnInvFormDelete_Click;
            pnlInvForm.Controls.Add(_btnInvFormDelete);

            InitializeStockMovementErrorTab(tabStockMovementErrors);

            // 의사면허 중복 관리를 기초 데이터 관리 서브 탭의 4번째 탭으로 추가
            _subTabDataManagement.TabPages.Add(_tabDoctorLicense);
        }

        private void InitializeStockMovementErrorTab(TabPage tabStockMovementErrors)
        {
            Panel pnlAuditTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 188,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            tabStockMovementErrors.Controls.Add(pnlAuditTop);

            Label lblDrugName = new Label { Text = "약품명", Location = new Point(16, 18), Size = new Size(60, 22), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtStockAuditDrugName = new TextBox { Text = "", Location = new Point(78, 15), Size = new Size(170, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _btnStockAuditDrugSearch = new Button
            {
                Text = "🔍 약품 검색",
                Location = new Point(256, 12),
                Size = new Size(118, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnStockAuditDrugSearch.FlatAppearance.BorderSize = 0;
            _btnStockAuditDrugSearch.Click += BtnStockAuditDrugSearch_Click;

            Label lblDrug = new Label { Text = "약품코드", Location = new Point(400, 18), Size = new Size(70, 22), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtStockAuditDrugCode = new TextBox { Text = "644913503", Location = new Point(472, 15), Size = new Size(130, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };

            Label lblUnit = new Label { Text = "입고 기준단위", Location = new Point(622, 18), Size = new Size(95, 22), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtStockAuditUnit = new TextBox { Text = "500", Location = new Point(720, 15), Size = new Size(80, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };

            Label lblMinQty = new Label { Text = "최소 처방총량", Location = new Point(818, 18), Size = new Size(95, 22), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtStockAuditMinQty = new TextBox { Text = "5", Location = new Point(916, 15), Size = new Size(80, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };

            _btnStockAuditRun = new Button
            {
                Text = "🔎 입출고 오류 검사",
                Location = new Point(1014, 12),
                Size = new Size(170, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnStockAuditRun.FlatAppearance.BorderSize = 0;
            _btnStockAuditRun.Click += BtnStockAuditRun_Click;

            Label lblHint = new Label
            {
                Text = "입고: 적수 × 수량이 기준단위 배수인지 검사 / 처방: 1회량 × 횟수 × 일수가 최소 처방총량 미만인지 검사",
                Location = new Point(16, 53),
                Size = new Size(900, 24),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };

            pnlAuditTop.Controls.Add(lblDrugName);
            pnlAuditTop.Controls.Add(_txtStockAuditDrugName);
            pnlAuditTop.Controls.Add(_btnStockAuditDrugSearch);
            pnlAuditTop.Controls.Add(lblDrug);
            pnlAuditTop.Controls.Add(_txtStockAuditDrugCode);
            pnlAuditTop.Controls.Add(lblUnit);
            pnlAuditTop.Controls.Add(_txtStockAuditUnit);
            pnlAuditTop.Controls.Add(lblMinQty);
            pnlAuditTop.Controls.Add(_txtStockAuditMinQty);
            pnlAuditTop.Controls.Add(_btnStockAuditRun);
            pnlAuditTop.Controls.Add(lblHint);

            _dgvStockAuditDrugSearch = new DataGridView
            {
                Location = new Point(16, 82),
                Size = new Size(560, 94),
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 28
            };
            _dgvStockAuditDrugSearch.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvStockAuditDrugSearch.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvStockAuditDrugSearch.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            _dgvStockAuditDrugSearch.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvStockAuditDrugSearch.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvStockAuditDrugSearch.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvStockAuditDrugSearch.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvStockAuditDrugSearch.CellClick += DgvStockAuditDrugSearch_CellClick;
            pnlAuditTop.Controls.Add(_dgvStockAuditDrugSearch);

            _txtStockAuditDrugInfo = new TextBox
            {
                Location = new Point(594, 82),
                Size = new Size(590, 94),
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlAuditTop.Controls.Add(_txtStockAuditDrugInfo);

            _dgvStockMovementErrors = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvStockMovementErrors.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvStockMovementErrors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvStockMovementErrors.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvStockMovementErrors.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvStockMovementErrors.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvStockMovementErrors.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvStockMovementErrors.DefaultCellStyle.SelectionForeColor = Color.White;
            tabStockMovementErrors.Controls.Add(_dgvStockMovementErrors);
            _dgvStockMovementErrors.BringToFront();

            InitializeStockAdjustmentRestoreTab();
        }

#region 재고보정 복구 (Stock Adjustment Restore)

        private void InitializeStockAdjustmentRestoreTab()
        {
            _tabStockAdjustmentRestore = new TabPage
            {
                Text = "재고보정 복구",
                BackColor = ColorBgMain
            };
            _subTabInventoryManagement.TabPages.Add(_tabStockAdjustmentRestore);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 78F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            _tabStockAdjustmentRestore.Controls.Add(layout);

            // Top Search & Action Panel
            Panel pnlTop = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            layout.Controls.Add(pnlTop, 0, 0);

            _btnStockAdjScan = new Button
            {
                Text = "재고보정 대조 검사 실행",
                Location = new Point(14, 8),
                Size = new Size(195, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnStockAdjScan.FlatAppearance.BorderSize = 0;
            _btnStockAdjScan.Click += BtnStockAdjScan_Click;
            pnlTop.Controls.Add(_btnStockAdjScan);

            _btnStockAdjRestoreSelected = new Button
            {
                Text = "선택 일자 복구",
                Location = new Point(217, 8),
                Size = new Size(135, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Enabled = false
            };
            _btnStockAdjRestoreSelected.FlatAppearance.BorderSize = 0;
            _btnStockAdjRestoreSelected.Click += BtnStockAdjRestoreSelected_Click;
            pnlTop.Controls.Add(_btnStockAdjRestoreSelected);

            _btnStockAdjRestoreAll = new Button
            {
                Text = "누락 전체 일괄 복구",
                Location = new Point(360, 8),
                Size = new Size(165, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Enabled = false
            };
            _btnStockAdjRestoreAll.FlatAppearance.BorderSize = 0;
            _btnStockAdjRestoreAll.Click += BtnStockAdjRestoreAll_Click;
            pnlTop.Controls.Add(_btnStockAdjRestoreAll);

            _btnStockAdjExportCsv = new Button
            {
                Text = "결과 CSV 저장",
                Location = new Point(533, 8),
                Size = new Size(125, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(51, 65, 85),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Enabled = false
            };
            _btnStockAdjExportCsv.FlatAppearance.BorderSize = 0;
            _btnStockAdjExportCsv.Click += BtnStockAdjExportCsv_Click;
            pnlTop.Controls.Add(_btnStockAdjExportCsv);

            _btnStockAdjAttachBackup = new Button
            {
                Text = "백업 DB 연결",
                Location = new Point(666, 8),
                Size = new Size(130, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(71, 85, 105),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnStockAdjAttachBackup.FlatAppearance.BorderSize = 0;
            _btnStockAdjAttachBackup.Click += BtnStockAdjAttachBackup_Click;
            pnlTop.Controls.Add(_btnStockAdjAttachBackup);

            _btnStockAdjDetachBackup = new Button
            {
                Text = "백업 DB 연결 해제",
                Location = new Point(804, 8),
                Size = new Size(155, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Enabled = false
            };
            _btnStockAdjDetachBackup.FlatAppearance.BorderSize = 0;
            _btnStockAdjDetachBackup.Click += BtnStockAdjDetachBackup_Click;
            pnlTop.Controls.Add(_btnStockAdjDetachBackup);

            _lblStockAdjBackupStatus = new Label
            {
                Text = "백업 DB 연결 상태를 확인하는 중입니다...",
                Location = new Point(14, 46),
                Size = new Size(1150, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Italic)
            };
            pnlTop.Controls.Add(_lblStockAdjBackupStatus);

            _tabStockAdjustmentRestore.Enter += (s, e) => RefreshStockAdjBackupStatus();

            // SplitContainer (Left: Summary list, Right: Detail list) - Safe Initialization
            _splitStockAdj = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Size = new Size(1100, 600),
                Panel1MinSize = 100,
                Panel2MinSize = 100,
                BackColor = ColorBorder
            };
            try
            {
                _splitStockAdj.SplitterDistance = Math.Max(100, Math.Min(700, _distStockAdj));
            }
            catch { }
            _splitStockAdj.SplitterMoved += (s, e) => _distStockAdj = _splitStockAdj.SplitterDistance;
            layout.Controls.Add(_splitStockAdj, 0, 1);

            // Left Panel (Summary)
            Panel pnlSummary = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(8)
            };
            _splitStockAdj.Panel1.Controls.Add(pnlSummary);

            TableLayoutPanel summaryLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = ColorBgCard
            };
            summaryLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            summaryLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            summaryLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            pnlSummary.Controls.Add(summaryLayout);

            Panel pnlSummaryHeader = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard
            };
            summaryLayout.Controls.Add(pnlSummaryHeader, 0, 0);

            Label lblSummaryTitle = new Label
            {
                Text = "재고보정 작업일자 목록",
                Location = new Point(4, 8),
                Size = new Size(165, 22),
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlSummaryHeader.Controls.Add(lblSummaryTitle);

            _lblStockAdjSummaryCount = new Label
            {
                Text = "",
                Location = new Point(172, 8),
                Size = new Size(135, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            pnlSummaryHeader.Controls.Add(_lblStockAdjSummaryCount);

            _btnStockAdjSelectMissingOnly = new Button
            {
                Text = "누락만 선택",
                Location = new Point(310, 4),
                Size = new Size(85, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _btnStockAdjSelectMissingOnly.FlatAppearance.BorderSize = 0;
            _btnStockAdjSelectMissingOnly.Click += BtnStockAdjSelectMissingOnly_Click;
            pnlSummaryHeader.Controls.Add(_btnStockAdjSelectMissingOnly);

            _btnStockAdjSelectAll = new Button
            {
                Text = "전체선택",
                Location = new Point(400, 4),
                Size = new Size(68, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            _btnStockAdjSelectAll.FlatAppearance.BorderSize = 0;
            _btnStockAdjSelectAll.Click += BtnStockAdjSelectAll_Click;
            pnlSummaryHeader.Controls.Add(_btnStockAdjSelectAll);

            _btnStockAdjDeselectAll = new Button
            {
                Text = "해제",
                Location = new Point(473, 4),
                Size = new Size(48, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            _btnStockAdjDeselectAll.FlatAppearance.BorderSize = 0;
            _btnStockAdjDeselectAll.Click += BtnStockAdjDeselectAll_Click;
            pnlSummaryHeader.Controls.Add(_btnStockAdjDeselectAll);

            _dgvStockAdjSummary = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 30
            };
            _dgvStockAdjSummary.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvStockAdjSummary.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvStockAdjSummary.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            _dgvStockAdjSummary.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvStockAdjSummary.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvStockAdjSummary.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvStockAdjSummary.DefaultCellStyle.SelectionForeColor = Color.White;
            _dgvStockAdjSummary.SelectionChanged += DgvStockAdjSummary_SelectionChanged;
            _dgvStockAdjSummary.CellContentClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 0)
                {
                    _dgvStockAdjSummary.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
            summaryLayout.Controls.Add(_dgvStockAdjSummary, 0, 1);

            // Right Panel (Detail)
            Panel pnlDetail = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(8)
            };
            _splitStockAdj.Panel2.Controls.Add(pnlDetail);

            TableLayoutPanel detailLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = ColorBgCard
            };
            detailLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            detailLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            pnlDetail.Controls.Add(detailLayout);

            _lblStockAdjDetailTitle = new Label
            {
                Text = "선택 일자 상세 약품 대조 내역 (좌측에서 보정일자를 선택하십시오)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            detailLayout.Controls.Add(_lblStockAdjDetailTitle, 0, 0);

            _dgvStockAdjDetail = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 30
            };
            _dgvStockAdjDetail.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvStockAdjDetail.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvStockAdjDetail.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            _dgvStockAdjDetail.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvStockAdjDetail.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvStockAdjDetail.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvStockAdjDetail.DefaultCellStyle.SelectionForeColor = Color.White;
            detailLayout.Controls.Add(_dgvStockAdjDetail, 0, 1);
        }

        private void RefreshStockAdjBackupStatus()
        {
            if (_lblStockAdjBackupStatus == null) return;

            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                _lblStockAdjBackupStatus.Text = "● 가상 데모 모드 (가상 백업 DB 연결됨: PM_MAIN_BACKUP_DEMO | 상태: 가상 시뮬레이션)";
                _lblStockAdjBackupStatus.ForeColor = ColorEmerald;
                if (_btnStockAdjDetachBackup != null) _btnStockAdjDetachBackup.Enabled = false;
                return;
            }

            bool readOnly;
            string databaseName = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (_btnStockAdjDetachBackup != null)
            {
                _btnStockAdjDetachBackup.Enabled = !string.IsNullOrEmpty(databaseName);
            }

            if (string.IsNullOrEmpty(databaseName))
            {
                _lblStockAdjBackupStatus.Text = "● 백업 DB 미연결: [백업 DB 연결]로 PM_MAIN.MDF 폴더를 선택하십시오.";
                _lblStockAdjBackupStatus.ForeColor = ColorWarning;
            }
            else
            {
                _lblStockAdjBackupStatus.Text = string.Format(
                    "● 연결된 백업 DB: {0}  |  상태: {1}",
                    databaseName,
                    readOnly ? "읽기 전용" : "주의 - 쓰기 가능");
                _lblStockAdjBackupStatus.ForeColor = readOnly ? ColorEmerald : ColorAlarm;
            }
        }

        private void BtnStockAdjAttachBackup_Click(object sender, EventArgs e)
        {
            BtnAttachPrescriptionBackup_Click(sender, e);
            RefreshStockAdjBackupStatus();
        }

        private void BtnStockAdjDetachBackup_Click(object sender, EventArgs e)
        {
            BtnDetachPrescriptionBackup_Click(sender, e);
            RefreshStockAdjBackupStatus();
        }

        private void BtnStockAdjScan_Click(object sender, EventArgs e)
        {
            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                ScanStockAdjustmentsDemo();
                return;
            }

            bool readOnly;
            string backupDb = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (string.IsNullOrEmpty(backupDb))
            {
                MessageBox.Show(
                    "재고보정 대조 검사를 실행하려면 먼저 [백업 DB 연결] 버튼을 통해 백업 데이터베이스를 연결해야 합니다.",
                    "백업 DB 필요",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            _lblStockAdjBackupStatus.Text = "운영 DB와 백업 DB 간의 재고보정 데이터를 전수 대조하는 중입니다...";
            Application.DoEvents();

            try
            {
                string connStr = BuildConnectionString(false);
                DataTable dt = new DataTable();

                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    string sql = string.Format(@"
WITH Bak AS (
    SELECT 
        h.PRES_DTIME,
        h.DRUG_SEQ,
        h.PRES_GUBUN,
        COUNT(d.MEDC_SEQ) AS Bak_Cnt04,
        (SELECT COUNT(*) FROM {0}.dbo.TBSID040_05 WITH (NOLOCK) WHERE DRUG_SEQ = h.DRUG_SEQ) AS Bak_Cnt05
    FROM {0}.dbo.TBSID040_03 h WITH (NOLOCK)
    LEFT JOIN {0}.dbo.TBSID040_04 d WITH (NOLOCK) ON h.DRUG_SEQ = d.DRUG_SEQ
    WHERE h.PRES_GUBUN = 'E' OR h.DRUG_SEQ LIKE '%099999'
    GROUP BY h.PRES_DTIME, h.DRUG_SEQ, h.PRES_GUBUN
),
Cur AS (
    SELECT 
        h.PRES_DTIME,
        h.DRUG_SEQ,
        h.PRES_GUBUN,
        COUNT(d.MEDC_SEQ) AS Cur_Cnt04,
        (SELECT COUNT(*) FROM PM_MAIN.dbo.TBSID040_05 WITH (NOLOCK) WHERE DRUG_SEQ = h.DRUG_SEQ) AS Cur_Cnt05
    FROM PM_MAIN.dbo.TBSID040_03 h WITH (NOLOCK)
    LEFT JOIN PM_MAIN.dbo.TBSID040_04 d WITH (NOLOCK) ON h.DRUG_SEQ = d.DRUG_SEQ
    WHERE h.PRES_GUBUN = 'E' OR h.DRUG_SEQ LIKE '%099999'
    GROUP BY h.PRES_DTIME, h.DRUG_SEQ, h.PRES_GUBUN
)
SELECT 
    CONVERT(bit, CASE WHEN (b.Bak_Cnt04 > 0 AND ISNULL(c.Cur_Cnt04, 0) = 0) OR (b.Bak_Cnt04 <> ISNULL(c.Cur_Cnt04, 0)) OR (c.DRUG_SEQ IS NULL) THEN 1 ELSE 0 END) AS [선택],
    ISNULL(b.PRES_DTIME, c.PRES_DTIME) AS [작업일자],
    ISNULL(b.DRUG_SEQ, c.DRUG_SEQ) AS [보정일련번호],
    ISNULL(b.Bak_Cnt04, 0) AS [백업품목수],
    ISNULL(c.Cur_Cnt04, 0) AS [운영품목수],
    ISNULL(b.Bak_Cnt05, 0) AS [백업조제수],
    ISNULL(c.Cur_Cnt05, 0) AS [운영조제수],
    CASE 
        WHEN b.DRUG_SEQ IS NOT NULL AND c.DRUG_SEQ IS NULL THEN '헤더 누락'
        WHEN b.Bak_Cnt04 > 0 AND ISNULL(c.Cur_Cnt04, 0) = 0 THEN '상세 유실 (0건)'
        WHEN b.Bak_Cnt04 <> ISNULL(c.Cur_Cnt04, 0) THEN '품목수 불일치'
        WHEN b.DRUG_SEQ IS NULL AND c.DRUG_SEQ IS NOT NULL THEN '운영 신규 작업'
        ELSE '정상 일치'
    END AS [대조상태],
    ISNULL(b.PRES_GUBUN, c.PRES_GUBUN) AS [구분]
FROM Bak b
FULL OUTER JOIN Cur c ON b.DRUG_SEQ = c.DRUG_SEQ
ORDER BY [작업일자] DESC, [보정일련번호] DESC;",
                        QuoteSqlName(backupDb));

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120;
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }

                _stockAdjSummaryDt = dt;
                _dgvStockAdjSummary.DataSource = dt;

                int missingCount = 0;
                foreach (DataRow r in dt.Rows)
                {
                    string st = r["대조상태"].ToString();
                    if (st != "정상 일치" && st != "운영 신규 작업") missingCount++;
                }

                _lblStockAdjSummaryCount.Text = string.Format("(총 {0}건 / 누락 {1}건)", dt.Rows.Count, missingCount);

                if (_dgvStockAdjSummary.Columns.Count > 0)
                {
                    _dgvStockAdjSummary.Columns["선택"].Width = 45;
                    _dgvStockAdjSummary.Columns["작업일자"].Width = 90;
                    _dgvStockAdjSummary.Columns["보정일련번호"].Width = 125;
                    _dgvStockAdjSummary.Columns["백업품목수"].Width = 75;
                    _dgvStockAdjSummary.Columns["운영품목수"].Width = 75;
                    _dgvStockAdjSummary.Columns["백업조제수"].Width = 75;
                    _dgvStockAdjSummary.Columns["운영조제수"].Width = 75;
                    _dgvStockAdjSummary.Columns["대조상태"].Width = 110;
                    _dgvStockAdjSummary.Columns["구분"].Width = 50;

                    for (int i = 1; i < _dgvStockAdjSummary.Columns.Count; i++)
                    {
                        _dgvStockAdjSummary.Columns[i].ReadOnly = true;
                    }
                }

                foreach (DataGridViewRow row in _dgvStockAdjSummary.Rows)
                {
                    string st = row.Cells["대조상태"].Value != null ? row.Cells["대조상태"].Value.ToString() : "";
                    if (st == "상세 유실 (0건)" || st == "헤더 누락")
                    {
                        row.DefaultCellStyle.ForeColor = Color.FromArgb(239, 68, 68);
                        row.DefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
                    }
                    else if (st == "품목수 불일치")
                    {
                        row.DefaultCellStyle.ForeColor = ColorWarning;
                        row.DefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
                    }
                }

                _btnStockAdjRestoreSelected.Enabled = (missingCount > 0);
                _btnStockAdjRestoreAll.Enabled = (missingCount > 0);
                _btnStockAdjExportCsv.Enabled = (dt.Rows.Count > 0);

                RefreshStockAdjBackupStatus();
                ShowToast(string.Format("검사 완료: 전체 {0}건 중 {1}건 누락/불일치 감지", dt.Rows.Count, missingCount), missingCount > 0 ? ColorWarning : ColorEmerald);
            }
            catch (Exception ex)
            {
                MessageBox.Show("재고보정 대조 검사 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RefreshStockAdjBackupStatus();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ScanStockAdjustmentsDemo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("선택", typeof(bool));
            dt.Columns.Add("작업일자", typeof(string));
            dt.Columns.Add("보정일련번호", typeof(string));
            dt.Columns.Add("백업품목수", typeof(int));
            dt.Columns.Add("운영품목수", typeof(int));
            dt.Columns.Add("백업조제수", typeof(int));
            dt.Columns.Add("운영조제수", typeof(int));
            dt.Columns.Add("대조상태", typeof(string));
            dt.Columns.Add("구분", typeof(string));

            string[] demoDates = new string[] {
                "20260317:1", "20250813:1", "20250731:5", "20231212:2", "20210907:1", "20210708:1",
                "20191014:1", "20190416:2", "20180928:2", "20180918:1", "20171114:8", "20170526:1",
                "20131108:3", "20130925:1", "20120129:1235", "20100330:2083"
            };

            foreach (string item in demoDates)
            {
                string[] parts = item.Split(':');
                string ymd = parts[0];
                int cnt = int.Parse(parts[1]);
                dt.Rows.Add(true, ymd, ymd + "099999", cnt, 0, cnt, 0, "상세 유실 (0건)", "E");
            }

            _stockAdjSummaryDt = dt;
            _dgvStockAdjSummary.DataSource = dt;
            _lblStockAdjSummaryCount.Text = string.Format("(총 {0}건 / 누락 {0}건)", dt.Rows.Count);

            if (_dgvStockAdjSummary.Columns.Count > 0)
            {
                _dgvStockAdjSummary.Columns["선택"].Width = 45;
                _dgvStockAdjSummary.Columns["작업일자"].Width = 90;
                _dgvStockAdjSummary.Columns["보정일련번호"].Width = 125;
                _dgvStockAdjSummary.Columns["백업품목수"].Width = 75;
                _dgvStockAdjSummary.Columns["운영품목수"].Width = 75;
                _dgvStockAdjSummary.Columns["백업조제수"].Width = 75;
                _dgvStockAdjSummary.Columns["운영조제수"].Width = 75;
                _dgvStockAdjSummary.Columns["대조상태"].Width = 110;
                _dgvStockAdjSummary.Columns["구분"].Width = 50;
            }

            foreach (DataGridViewRow row in _dgvStockAdjSummary.Rows)
            {
                row.DefaultCellStyle.ForeColor = Color.FromArgb(239, 68, 68);
                row.DefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            }

            _btnStockAdjRestoreSelected.Enabled = true;
            _btnStockAdjRestoreAll.Enabled = true;
            _btnStockAdjExportCsv.Enabled = true;
            ShowToast("[데모] 16건의 누락 보정일자가 검출되었습니다.", ColorWarning);
        }

        private void DgvStockAdjSummary_SelectionChanged(object sender, EventArgs e)
        {
            if (_dgvStockAdjSummary.CurrentRow == null) return;

            DataGridViewRow row = _dgvStockAdjSummary.CurrentRow;
            string drugSeq = row.Cells["보정일련번호"].Value != null ? row.Cells["보정일련번호"].Value.ToString() : "";
            string workDate = row.Cells["작업일자"].Value != null ? row.Cells["작업일자"].Value.ToString() : "";
            int bakCount = row.Cells["백업품목수"].Value != null ? Convert.ToInt32(row.Cells["백업품목수"].Value) : 0;
            int curCount = row.Cells["운영품목수"].Value != null ? Convert.ToInt32(row.Cells["운영품목수"].Value) : 0;

            _lblStockAdjDetailTitle.Text = string.Format("선택 작업: {0} ({1}) - 백업 {2:N0}건 / 운영 {3:N0}건", workDate, drugSeq, bakCount, curCount);

            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                LoadStockAdjDetailDemo(drugSeq, workDate);
                return;
            }

            bool readOnly;
            string backupDb = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (string.IsNullOrEmpty(backupDb) || string.IsNullOrEmpty(drugSeq))
            {
                _dgvStockAdjDetail.DataSource = null;
                return;
            }

            try
            {
                string connStr = BuildConnectionString(false);
                DataTable dt = new DataTable();

                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string sql = string.Format(@"
SELECT TOP 500
    b.DRUG_CODE AS [약품코드],
    ISNULL(m.ARTCNM, '') AS [약품명],
    ISNULL(m.MNF_CO_NM, '') AS [제조회사],
    b.DD_MQTY AS [입력수량],
    b.MDCN_MQTY AS [재고차이],
    ISNULL(s8.MDCN_MQTY, 0) AS [백업보정기준(08)],
    CASE 
        WHEN c.DRUG_CODE IS NULL THEN '운영 누락'
        ELSE '운영 존재'
    END AS [상태]
FROM {0}.dbo.TBSID040_04 b WITH (NOLOCK)
LEFT JOIN PM_MAIN.dbo.TBSID040_04 c WITH (NOLOCK) ON b.DRUG_SEQ = c.DRUG_SEQ AND b.MEDC_SEQ = c.MEDC_SEQ
LEFT JOIN {0}.dbo.TBSIM040_01 m WITH (NOLOCK) ON b.DRUG_CODE = m.DRUG_CODE
LEFT JOIN {0}.dbo.TBSIM040_08 s8 WITH (NOLOCK) ON b.DRUG_CODE = s8.DRUG_CODE AND s8.PRES_DTIME = b.PRES_DTIME
WHERE b.DRUG_SEQ = @seq
ORDER BY b.MEDC_SEQ ASC;",
                        QuoteSqlName(backupDb));

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@seq", drugSeq);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }

                _dgvStockAdjDetail.DataSource = dt;
                if (_dgvStockAdjDetail.Columns.Count > 0)
                {
                    _dgvStockAdjDetail.Columns["약품코드"].Width = 90;
                    _dgvStockAdjDetail.Columns["약품명"].Width = 180;
                    _dgvStockAdjDetail.Columns["제조회사"].Width = 110;
                    _dgvStockAdjDetail.Columns["입력수량"].Width = 75;
                    _dgvStockAdjDetail.Columns["재고차이"].Width = 75;
                    _dgvStockAdjDetail.Columns["백업보정기준(08)"].Width = 90;
                    _dgvStockAdjDetail.Columns["상태"].Width = 80;
                }

                foreach (DataGridViewRow r in _dgvStockAdjDetail.Rows)
                {
                    string st = r.Cells["상태"].Value != null ? r.Cells["상태"].Value.ToString() : "";
                    if (st == "운영 누락")
                    {
                        r.DefaultCellStyle.ForeColor = Color.FromArgb(239, 68, 68);
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStockAdjDetailTitle.Text = "상세 조회 실패: " + ex.Message;
            }
        }

        private void LoadStockAdjDetailDemo(string drugSeq, string workDate)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("약품코드", typeof(string));
            dt.Columns.Add("약품명", typeof(string));
            dt.Columns.Add("제조회사", typeof(string));
            dt.Columns.Add("입력수량", typeof(decimal));
            dt.Columns.Add("재고차이", typeof(decimal));
            dt.Columns.Add("백업보정기준(08)", typeof(decimal));
            dt.Columns.Add("상태", typeof(string));

            dt.Rows.Add("668900990", "자니딥정 10mg", "LG화학", 377, 13602, 377, "운영 누락");
            dt.Rows.Add("655500310", "자누비아정 100mg", "한국MSD", 0, -235, 0, "운영 누락");
            dt.Rows.Add("643309540", "자누비아정 50mg", "한국MSD", 121, 235, 121, "운영 누락");

            _dgvStockAdjDetail.DataSource = dt;
        }

        private void BtnStockAdjSelectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in _dgvStockAdjSummary.Rows)
            {
                r.Cells["선택"].Value = true;
            }
        }

        private void BtnStockAdjDeselectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in _dgvStockAdjSummary.Rows)
            {
                r.Cells["선택"].Value = false;
            }
        }

        private void BtnStockAdjSelectMissingOnly_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in _dgvStockAdjSummary.Rows)
            {
                string st = r.Cells["대조상태"].Value != null ? r.Cells["대조상태"].Value.ToString() : "";
                r.Cells["선택"].Value = (st != "정상 일치" && st != "운영 신규 작업");
            }
        }

        private void BtnStockAdjRestoreSelected_Click(object sender, EventArgs e)
        {
            List<string> selectedSeqs = new List<string>();
            foreach (DataGridViewRow r in _dgvStockAdjSummary.Rows)
            {
                bool isChecked = Convert.ToBoolean(r.Cells["선택"].Value);
                if (isChecked)
                {
                    string seq = r.Cells["보정일련번호"].Value != null ? r.Cells["보정일련번호"].Value.ToString() : "";
                    if (!string.IsNullOrEmpty(seq)) selectedSeqs.Add(seq);
                }
            }

            if (selectedSeqs.Count == 0)
            {
                MessageBox.Show("복구할 재고보정 작업 일자를 목록에서 선택(체크)해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ExecuteStockAdjustmentRestore(selectedSeqs);
        }

        private void BtnStockAdjRestoreAll_Click(object sender, EventArgs e)
        {
            List<string> missingSeqs = new List<string>();
            foreach (DataGridViewRow r in _dgvStockAdjSummary.Rows)
            {
                string st = r.Cells["대조상태"].Value != null ? r.Cells["대조상태"].Value.ToString() : "";
                if (st != "정상 일치" && st != "운영 신규 작업")
                {
                    string seq = r.Cells["보정일련번호"].Value != null ? r.Cells["보정일련번호"].Value.ToString() : "";
                    if (!string.IsNullOrEmpty(seq)) missingSeqs.Add(seq);
                }
            }

            if (missingSeqs.Count == 0)
            {
                MessageBox.Show("복구 대상인 누락/불일치 재고보정 건이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExecuteStockAdjustmentRestore(missingSeqs);
        }

        private void ExecuteStockAdjustmentRestore(List<string> drugSeqs)
        {
            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                MessageBox.Show(
                    string.Format("[데모] 선택된 {0}개 재고보정 일자(가상 3,348건)의 복구가 성공적으로 시뮬레이션 완료되었습니다.", drugSeqs.Count),
                    "복구 완료 (데모)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                ShowToast("재고보정 복구 완료 (데모)", ColorEmerald);
                return;
            }

            bool readOnly;
            string backupDb = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (string.IsNullOrEmpty(backupDb))
            {
                MessageBox.Show("백업 데이터베이스가 연결되어 있지 않습니다. [백업 DB 연결]을 먼저 진행하십시오.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DialogResult dr = MessageBox.Show(
                string.Format(
                    "선택한 총 {0}개 재고보정 작업(DRUG_SEQ)의 상세 약품 및 기준 데이터를 백업 DB [{1}]에서 운영 DB로 안전하게 복원하시겠습니까?\n\n" +
                    "※ 복원 대상 테이블:\n" +
                    " - TBSID040_03 (보정 마스터 헤더)\n" +
                    " - TBSID040_04 (상세 보정 약품 품목)\n" +
                    " - TBSID040_05 (조제 상세 내역)\n" +
                    " - TBSIM040_08 (약품별 실사 재고보정 기준)\n\n" +
                    "※ 기존의 다른 정상 데이터는 전혀 영향을 받지 않으며 트랜잭션으로 안전하게 보호됩니다.",
                    drugSeqs.Count, backupDb),
                "재고보정 데이터 복구 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dr != DialogResult.Yes) return;

            this.Cursor = Cursors.WaitCursor;
            try
            {
                int restored03 = 0, restored04 = 0, restored05 = 0, restored08 = 0;
                string connStr = BuildConnectionString(false);

                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlTransaction trans = conn.BeginTransaction())
                    {
                        try
                        {
                            foreach (string seq in drugSeqs)
                            {
                                string workDate = seq.Length >= 8 ? seq.Substring(0, 8) : "";

                                // 1. TBSID040_03 헤더 복원 (누락된 경우에만)
                                string sql03 = string.Format(@"
                                    INSERT INTO PM_MAIN.dbo.TBSID040_03
                                    SELECT b.* 
                                    FROM {0}.dbo.TBSID040_03 b WITH (NOLOCK)
                                    WHERE b.DRUG_SEQ = @seq
                                      AND NOT EXISTS (
                                          SELECT 1 FROM PM_MAIN.dbo.TBSID040_03 c WITH (NOLOCK) 
                                          WHERE c.DRUG_SEQ = b.DRUG_SEQ
                                      );", QuoteSqlName(backupDb));

                                using (SqlCommand cmd03 = new SqlCommand(sql03, conn, trans))
                                {
                                    cmd03.Parameters.AddWithValue("@seq", seq);
                                    restored03 += cmd03.ExecuteNonQuery();
                                }

                                // 2. TBSID040_04 상세 약품 복원
                                string sql04 = string.Format(@"
                                    INSERT INTO PM_MAIN.dbo.TBSID040_04
                                    SELECT b.* 
                                    FROM {0}.dbo.TBSID040_04 b WITH (NOLOCK)
                                    WHERE b.DRUG_SEQ = @seq
                                      AND NOT EXISTS (
                                          SELECT 1 FROM PM_MAIN.dbo.TBSID040_04 c WITH (NOLOCK) 
                                          WHERE c.DRUG_SEQ = b.DRUG_SEQ AND c.MEDC_SEQ = b.MEDC_SEQ
                                      );", QuoteSqlName(backupDb));

                                using (SqlCommand cmd04 = new SqlCommand(sql04, conn, trans))
                                {
                                    cmd04.Parameters.AddWithValue("@seq", seq);
                                    restored04 += cmd04.ExecuteNonQuery();
                                }

                                // 3. TBSID040_05 조제 상세 복원
                                string sql05 = string.Format(@"
                                    INSERT INTO PM_MAIN.dbo.TBSID040_05
                                    SELECT b.* 
                                    FROM {0}.dbo.TBSID040_05 b WITH (NOLOCK)
                                    WHERE b.DRUG_SEQ = @seq
                                      AND NOT EXISTS (
                                          SELECT 1 FROM PM_MAIN.dbo.TBSID040_05 c WITH (NOLOCK) 
                                          WHERE c.DRUG_SEQ = b.DRUG_SEQ AND c.MEDC_SEQ = b.MEDC_SEQ
                                      );", QuoteSqlName(backupDb));

                                using (SqlCommand cmd05 = new SqlCommand(sql05, conn, trans))
                                {
                                    cmd05.Parameters.AddWithValue("@seq", seq);
                                    restored05 += cmd05.ExecuteNonQuery();
                                }

                                // 4. TBSIM040_08 약품별 보정 기준치 복원
                                if (!string.IsNullOrEmpty(workDate))
                                {
                                    string sql08 = string.Format(@"
                                        INSERT INTO PM_MAIN.dbo.TBSIM040_08
                                        SELECT b.* 
                                        FROM {0}.dbo.TBSIM040_08 b WITH (NOLOCK)
                                        WHERE b.PRES_DTIME = @presDate
                                          AND NOT EXISTS (
                                              SELECT 1 FROM PM_MAIN.dbo.TBSIM040_08 c WITH (NOLOCK) 
                                              WHERE c.DRUG_CODE = b.DRUG_CODE AND c.PRES_DTIME = b.PRES_DTIME
                                          );", QuoteSqlName(backupDb));

                                    using (SqlCommand cmd08 = new SqlCommand(sql08, conn, trans))
                                    {
                                        cmd08.Parameters.AddWithValue("@presDate", workDate);
                                        restored08 += cmd08.ExecuteNonQuery();
                                    }
                                }
                            }

                            trans.Commit();
                        }
                        catch (Exception)
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }

                ShowToast(string.Format("재고보정 복구 성공: 총 {0}개 약품 상세 복원 완료", restored04), ColorEmerald);
                MessageBox.Show(
                    string.Format(
                        "재고보정 데이터 복구가 성공적으로 완료되었습니다.\n\n" +
                        "■ 복구 결과 요약:\n" +
                        " - 처리 대상 작업일자: {0}건\n" +
                        " - 복구된 헤더(TBSID040_03): {1}건\n" +
                        " - 복구된 약품상세(TBSID040_04): {2:N0}건\n" +
                        " - 복구된 조제상세(TBSID040_05): {3:N0}건\n" +
                        " - 복구된 보정기준(TBSIM040_08): {4:N0}건\n\n" +
                        "※ 유팜(PM+)의 [재고보정 / 보정현황] 화면에서 정상 표시되는지 확인하십시오.",
                        drugSeqs.Count, restored03, restored04, restored05, restored08),
                    "재고보정 복구 완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Re-scan automatically
                BtnStockAdjScan_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("재고보정 복구 작업 중 오류가 발생하여 모든 변경 사항이 롤백되었습니다:\n\n" + ex.Message, "복구 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void BtnStockAdjExportCsv_Click(object sender, EventArgs e)
        {
            if (_stockAdjSummaryDt == null || _stockAdjSummaryDt.Rows.Count == 0)
            {
                MessageBox.Show("저장할 검사 결과가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 파일 (*.csv)|*.csv";
                dialog.FileName = string.Format("재고보정_대조검사결과_{0}.csv", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("작업일자,보정일련번호,백업품목수,운영품목수,백업조제수,운영조제수,대조상태,구분");

                        foreach (DataRow row in _stockAdjSummaryDt.Rows)
                        {
                            sb.AppendLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                row["작업일자"],
                                row["보정일련번호"],
                                row["백업품목수"],
                                row["운영품목수"],
                                row["백업조제수"],
                                row["운영조제수"],
                                row["대조상태"],
                                row["구분"]));
                        }

                        File.WriteAllText(dialog.FileName, sb.ToString(), Encoding.UTF8);
                        ShowToast("CSV 파일이 성공적으로 저장되었습니다.", ColorEmerald);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("CSV 저장 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion

        private void InitializeClaimComparisonTab()
        {
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            _tabClaimComparison.Controls.Add(layout);

            Panel pnlSearch = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            layout.Controls.Add(pnlSearch, 0, 0);

            Label lblMonth = new Label
            {
                Text = "조회월",
                Location = new Point(18, 25),
                Size = new Size(52, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _dtpClaimComparisonMonth = new DateTimePicker
            {
                Location = new Point(76, 21),
                Size = new Size(130, 27),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy년 MM월",
                ShowUpDown = true,
                Value = new DateTime(DateTime.Today.AddMonths(-1).Year, DateTime.Today.AddMonths(-1).Month, 1),
                CalendarForeColor = ColorTextMain,
                CalendarMonthBackground = ColorBgMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            Label lblInsuranceType = new Label
            {
                Text = "구분",
                Location = new Point(224, 25),
                Size = new Size(38, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _cmbClaimComparisonType = new ComboBox
            {
                Location = new Point(266, 21),
                Size = new Size(145, 27),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            _cmbClaimComparisonType.Items.AddRange(new object[] { "전체", "건강보험", "의료급여(보호)", "보훈" });
            _cmbClaimComparisonType.SelectedIndex = 0;
            _chkClaimComparisonExcludeZero = new CheckBox
            {
                Text = "청구액 0원 제외",
                Location = new Point(428, 22),
                Size = new Size(150, 27),
                Checked = true,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnClaimComparisonSearch = new Button
            {
                Text = "🔍 월별 비교 조회",
                Location = new Point(590, 18),
                Size = new Size(150, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnClaimComparisonSearch.FlatAppearance.BorderSize = 0;
            _btnClaimComparisonSearch.Click += BtnClaimComparisonSearch_Click;

            _btnClaimComparisonExport = new Button
            {
                Text = "💾 CSV 저장",
                Location = new Point(750, 18),
                Size = new Size(125, 34),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(8, 145, 178),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Enabled = false
            };
            _btnClaimComparisonExport.FlatAppearance.BorderSize = 0;
            _btnClaimComparisonExport.Click += BtnClaimComparisonExport_Click;

            pnlSearch.Controls.Add(lblMonth);
            pnlSearch.Controls.Add(_dtpClaimComparisonMonth);
            pnlSearch.Controls.Add(lblInsuranceType);
            pnlSearch.Controls.Add(_cmbClaimComparisonType);
            pnlSearch.Controls.Add(_chkClaimComparisonExcludeZero);
            pnlSearch.Controls.Add(_btnClaimComparisonSearch);
            pnlSearch.Controls.Add(_btnClaimComparisonExport);

            _dgvClaimComparison = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32,
                ScrollBars = ScrollBars.Both
            };
            _dgvClaimComparison.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvClaimComparison.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvClaimComparison.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvClaimComparison.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvClaimComparison.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvClaimComparison.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvClaimComparison.DefaultCellStyle.SelectionForeColor = Color.White;
            layout.Controls.Add(_dgvClaimComparison, 0, 1);

            _lblClaimComparisonSummary = new Label
            {
                Text = "조회 전",
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                ForeColor = ColorTextSec,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            layout.Controls.Add(_lblClaimComparisonSummary, 0, 2);
        }

        private void BtnClaimComparisonSearch_Click(object sender, EventArgs e)
        {
            DateTime monthStart = new DateTime(_dtpClaimComparisonMonth.Value.Year, _dtpClaimComparisonMonth.Value.Month, 1);
            DateTime monthEnd = monthStart.AddMonths(1);
            string selectedInsuranceType = Convert.ToString(_cmbClaimComparisonType.SelectedItem);
            DataTable result = new DataTable();
            DataTable summary = new DataTable();

            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (_chkDemoMode.Checked)
                {
                    result.Columns.Add("조제일자");
                    result.Columns.Add("환자명");
                    result.Columns.Add("차트번호");
                    result.Columns.Add("처방번호");
                    result.Columns.Add("구분");
                    result.Columns.Add("상태");
                    result.Columns.Add("총약제비", typeof(decimal));
                    result.Columns.Add("청구액", typeof(decimal));
                    result.Columns.Add("본인부담", typeof(decimal));
                    result.Columns.Add("입금액", typeof(decimal));
                    result.Columns.Add("대체청구번호");
                    result.Columns.Add("비교결과");
                    result.Rows.Add(monthStart.ToString("yyyyMMdd"), "홍길동", "0000000001", monthStart.ToString("yyyyMM") + "000001", "보훈", "5", 305720m, 305724m, 0m, 0m, "", "청구 누락 후보");
                    result.Rows.Add(monthStart.ToString("yyyyMMdd"), "김환자", "0000000002", monthStart.ToString("yyyyMM") + "000002", "의료급여(보호)", "5", 255190m, 242600m, 500m, 12590m, monthStart.ToString("yyyyMM") + "009999", "처방번호 불일치 확인");

                    if (!string.IsNullOrEmpty(selectedInsuranceType) && selectedInsuranceType != "전체")
                    {
                        for (int i = result.Rows.Count - 1; i >= 0; i--)
                        {
                            if (Convert.ToString(result.Rows[i]["구분"]) != selectedInsuranceType) result.Rows.RemoveAt(i);
                        }
                    }

                    summary.Columns.Add("전체", typeof(int));
                    summary.Columns.Add("직접청구", typeof(int));
                    summary.Columns.Add("번호불일치", typeof(int));
                    summary.Columns.Add("누락", typeof(int));
                    summary.Columns.Add("청구액0", typeof(int));
                    summary.Rows.Add(result.Rows.Count + 10, 10, result.Rows.Count > 1 ? 1 : 0, result.Rows.Count > 0 ? 1 : 0, 0);
                }
                else
                {
                    string sql = @"
WITH missing AS
(
    SELECT r.*
    FROM dbo.TBSID040_03 r WITH (NOLOCK)
    WHERE r.PRES_DTIME >= @monthStart
      AND r.PRES_DTIME < @monthEnd
      AND NOT EXISTS
      (
          SELECT 1
          FROM dbo.TBSIB_H024_1 h WITH (NOLOCK)
          WHERE h.DRUG_SEQ = r.DRUG_SEQ
      )
), calc AS
(
    SELECT m.*,
        CASE
            WHEN m.MPRE_TYPE = '4' AND m.MPRE_TYPE_GUBUN = '7' THEN N'보훈'
            WHEN m.MPRE_TYPE = '1' THEN N'의료급여(보호)'
            WHEN m.MPRE_TYPE = '0' THEN N'건강보험'
            ELSE N'기타(' + ISNULL(m.MPRE_TYPE, N'NULL') + N'/' + ISNULL(m.MPRE_TYPE_GUBUN, N'NULL') + N')'
        END AS INSURANCE_TYPE,
        CASE
            WHEN m.MPRE_TYPE = '4' AND m.MPRE_TYPE_GUBUN = '7'
                THEN ISNULL(m.VET_INS_PRICE, 0)
                   + ISNULL(m.EXP_UNDER_VET_INS_PRICE, 0)
                   + ISNULL(m.EXP_UNDER_INS_PRICE, 0)
            ELSE ISNULL(m.INS_PRICE, 0) + ISNULL(m.EXP_UNDER_INS_PRICE, 0)
        END AS CLAIM_AMOUNT,
        ISNULL(m.EXE_PRICE, 0) AS COPAY_AMOUNT,
        ISNULL(m.EXE_PRICE, 0)
          + ISNULL(m.EXP_EXE_PRICE, 0)
          + ISNULL(m.EXP_NON_EXE_PRICE, 0)
          + ISNULL(m.EXP_NON_PRICE, 0)
          + ISNULL(m.EXP_UNDER_EXE_PRICE, 0)
          + ISNULL(m.NON_PREP_PRICE, 0)
          + ISNULL(m.NON_DRUG_PRICE, 0) AS RECEIPT_AMOUNT
    FROM missing m
), compared AS
(
    SELECT c.*, alt.DRUG_SEQ AS ALT_CLAIM_DRUG_SEQ
    FROM calc c
    OUTER APPLY
    (
        SELECT TOP (1) h.DRUG_SEQ
        FROM dbo.TBSIB_H024_1 h WITH (NOLOCK)
        WHERE NULLIF(LTRIM(RTRIM(c.MPRSC_GRANT_NO)), '') IS NOT NULL
          AND h.MPRSC_GRANT_NO = c.MPRSC_GRANT_NO
        ORDER BY h.DRUG_SEQ
    ) alt
)
SELECT
    c.PRES_DTIME AS [조제일자],
    c.PAT_NM AS [환자명],
    c.CHRTNO AS [차트번호],
    c.DRUG_SEQ AS [처방번호],
    c.INSURANCE_TYPE AS [구분],
    c.PRES_PRGRS_STATE AS [상태],
    CASE
        WHEN c.INSURANCE_TYPE = N'보훈' THEN
            ROUND(CONVERT(decimal(18,2),
                ISNULL(c.VET_TOT_PRICE, 0)
              + ISNULL(c.EXP_UNDER_TOT_PRICE, 0)
              + ISNULL(c.EXP_EXE_PRICE, 0)
              + ISNULL(c.EXP_NON_EXE_PRICE, 0)
              + ISNULL(c.EXP_NON_PRICE, 0)
              + ISNULL(c.EXP_UNDER_EXE_PRICE, 0)
              + ISNULL(c.NON_PREP_PRICE, 0)
              + ISNULL(c.NON_DRUG_PRICE, 0)), -1)
        ELSE CONVERT(decimal(18,2), c.CLAIM_AMOUNT + c.RECEIPT_AMOUNT)
    END AS [총약제비],
    CONVERT(decimal(18,2), c.CLAIM_AMOUNT) AS [청구액],
    CONVERT(decimal(18,2), c.COPAY_AMOUNT) AS [본인부담],
    CONVERT(decimal(18,2), c.RECEIPT_AMOUNT) AS [입금액],
    ISNULL(c.ALT_CLAIM_DRUG_SEQ, '') AS [대체청구번호],
    CASE WHEN c.ALT_CLAIM_DRUG_SEQ IS NULL
         THEN N'청구 누락 후보'
         WHEN LEFT(c.DRUG_SEQ, 8) <> LEFT(c.PRES_DTIME, 8)
         THEN N'처방번호 날짜 불일치'
         ELSE N'처방번호 불일치 확인'
    END AS [비교결과]
FROM compared c
WHERE (@excludeZero = 0 OR c.CLAIM_AMOUNT > 0)
  AND (@insuranceType = N'전체' OR c.INSURANCE_TYPE = @insuranceType)
ORDER BY
    CASE WHEN c.ALT_CLAIM_DRUG_SEQ IS NULL THEN 0 ELSE 1 END,
    c.PRES_DTIME,
    c.DRUG_SEQ;";

                    string summarySql = @"
WITH rx AS
(
    SELECT r.*,
        CASE
            WHEN r.MPRE_TYPE = '4' AND r.MPRE_TYPE_GUBUN = '7' THEN N'보훈'
            WHEN r.MPRE_TYPE = '1' THEN N'의료급여(보호)'
            WHEN r.MPRE_TYPE = '0' THEN N'건강보험'
            ELSE N'기타(' + ISNULL(r.MPRE_TYPE, N'NULL') + N'/' + ISNULL(r.MPRE_TYPE_GUBUN, N'NULL') + N')'
        END AS INSURANCE_TYPE,
        CASE
            WHEN r.MPRE_TYPE = '4' AND r.MPRE_TYPE_GUBUN = '7'
                THEN ISNULL(r.VET_INS_PRICE, 0)
                   + ISNULL(r.EXP_UNDER_VET_INS_PRICE, 0)
                   + ISNULL(r.EXP_UNDER_INS_PRICE, 0)
            ELSE ISNULL(r.INS_PRICE, 0) + ISNULL(r.EXP_UNDER_INS_PRICE, 0)
        END AS CLAIM_AMOUNT
    FROM dbo.TBSID040_03 r WITH (NOLOCK)
    WHERE r.PRES_DTIME >= @monthStart
      AND r.PRES_DTIME < @monthEnd
), tagged AS
(
    SELECT rx.*,
        CASE WHEN EXISTS
        (
            SELECT 1 FROM dbo.TBSIB_H024_1 h WITH (NOLOCK)
            WHERE h.DRUG_SEQ = rx.DRUG_SEQ
        ) THEN 1 ELSE 0 END AS DIRECT_CLAIM,
        alt.DRUG_SEQ AS ALT_CLAIM_DRUG_SEQ
    FROM rx
    OUTER APPLY
    (
        SELECT TOP (1) h.DRUG_SEQ
        FROM dbo.TBSIB_H024_1 h WITH (NOLOCK)
        WHERE NULLIF(LTRIM(RTRIM(rx.MPRSC_GRANT_NO)), '') IS NOT NULL
          AND h.MPRSC_GRANT_NO = rx.MPRSC_GRANT_NO
        ORDER BY h.DRUG_SEQ
    ) alt
)
SELECT
    COUNT(*) AS [전체],
    SUM(CASE WHEN CLAIM_AMOUNT > 0 AND DIRECT_CLAIM = 1 THEN 1 ELSE 0 END) AS [직접청구],
    SUM(CASE WHEN CLAIM_AMOUNT > 0 AND DIRECT_CLAIM = 0 AND ALT_CLAIM_DRUG_SEQ IS NOT NULL THEN 1 ELSE 0 END) AS [번호불일치],
    SUM(CASE WHEN CLAIM_AMOUNT > 0 AND DIRECT_CLAIM = 0 AND ALT_CLAIM_DRUG_SEQ IS NULL THEN 1 ELSE 0 END) AS [누락],
    SUM(CASE WHEN CLAIM_AMOUNT = 0 THEN 1 ELSE 0 END) AS [청구액0]
FROM tagged
WHERE (@insuranceType = N'전체' OR INSURANCE_TYPE = @insuranceType);";

                    using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    using (SqlCommand summaryCmd = new SqlCommand(summarySql, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    using (SqlDataAdapter summaryAdapter = new SqlDataAdapter(summaryCmd))
                    {
                        cmd.CommandTimeout = 120;
                        cmd.Parameters.Add("@monthStart", SqlDbType.NVarChar, 8).Value = monthStart.ToString("yyyyMMdd");
                        cmd.Parameters.Add("@monthEnd", SqlDbType.NVarChar, 8).Value = monthEnd.ToString("yyyyMMdd");
                        cmd.Parameters.Add("@excludeZero", SqlDbType.Bit).Value = _chkClaimComparisonExcludeZero.Checked;
                        cmd.Parameters.Add("@insuranceType", SqlDbType.NVarChar, 30).Value = string.IsNullOrEmpty(selectedInsuranceType) ? "전체" : selectedInsuranceType;

                        summaryCmd.CommandTimeout = 120;
                        summaryCmd.Parameters.Add("@monthStart", SqlDbType.NVarChar, 8).Value = monthStart.ToString("yyyyMMdd");
                        summaryCmd.Parameters.Add("@monthEnd", SqlDbType.NVarChar, 8).Value = monthEnd.ToString("yyyyMMdd");
                        summaryCmd.Parameters.Add("@insuranceType", SqlDbType.NVarChar, 30).Value = string.IsNullOrEmpty(selectedInsuranceType) ? "전체" : selectedInsuranceType;

                        conn.Open();
                        adapter.Fill(result);
                        summaryAdapter.Fill(summary);
                    }
                }

                _dgvClaimComparison.DataSource = result;
                ApplyContentSizedColumns(_dgvClaimComparison);
                FormatClaimComparisonGrid();
                UpdateClaimComparisonSummary(result, summary, selectedInsuranceType);
                _btnClaimComparisonExport.Enabled = result.Rows.Count > 0;
                ShowToast(string.Format("{0:yyyy년 MM월} 청구 비교 {1}건 조회", monthStart, result.Rows.Count), ColorEmerald);
            }
            catch (Exception ex)
            {
                _btnClaimComparisonExport.Enabled = false;
                MessageBox.Show("청구 누락 비교 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void FormatClaimComparisonGrid()
        {
            if (_dgvClaimComparison == null) return;

            string[] amountColumns = { "총약제비", "청구액", "본인부담", "입금액" };
            foreach (string columnName in amountColumns)
            {
                if (!_dgvClaimComparison.Columns.Contains(columnName)) continue;
                DataGridViewColumn column = _dgvClaimComparison.Columns[columnName];
                column.DefaultCellStyle.Format = "N0";
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void UpdateClaimComparisonSummary(DataTable result, DataTable summary, string insuranceType)
        {
            decimal claimTotal = 0m;

            foreach (DataRow row in result.Rows)
            {
                if (row["청구액"] != DBNull.Value)
                {
                    claimTotal += Convert.ToDecimal(row["청구액"]);
                }
            }

            if (summary != null && summary.Rows.Count > 0)
            {
                DataRow row = summary.Rows[0];
                int total = row["전체"] == DBNull.Value ? 0 : Convert.ToInt32(row["전체"]);
                int direct = row["직접청구"] == DBNull.Value ? 0 : Convert.ToInt32(row["직접청구"]);
                int mismatch = row["번호불일치"] == DBNull.Value ? 0 : Convert.ToInt32(row["번호불일치"]);
                int missing = row["누락"] == DBNull.Value ? 0 : Convert.ToInt32(row["누락"]);
                int zero = row["청구액0"] == DBNull.Value ? 0 : Convert.ToInt32(row["청구액0"]);

                _lblClaimComparisonSummary.Text = string.Format(
                    "{0} 전체 {1:N0}건 = 직접 청구 {2:N0} + 처방번호 불일치 {3:N0} + 누락 후보 {4:N0} + 청구액 0원 {5:N0}  |  목록 {6:N0}건 / 청구액 {7:N0}원",
                    string.IsNullOrEmpty(insuranceType) ? "전체" : insuranceType,
                    total, direct, mismatch, missing, zero, result.Rows.Count, claimTotal);
            }
            else
            {
                _lblClaimComparisonSummary.Text = string.Format("목록 {0:N0}건  |  청구액 {1:N0}원", result.Rows.Count, claimTotal);
            }
        }

        private void BtnClaimComparisonExport_Click(object sender, EventArgs e)
        {
            if (_dgvClaimComparison == null || _dgvClaimComparison.Rows.Count == 0) return;

            DateTime selectedMonth = _dtpClaimComparisonMonth.Value;
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 파일 (*.csv)|*.csv";
                dialog.FileName = string.Format("{0:yyyy-MM}_조제청구비교.csv", selectedMonth);
                dialog.Title = "청구 비교 결과 저장";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    using (StreamWriter writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true)))
                    {
                        List<DataGridViewColumn> visibleColumns = _dgvClaimComparison.Columns
                            .Cast<DataGridViewColumn>()
                            .Where(c => c.Visible)
                            .OrderBy(c => c.DisplayIndex)
                            .ToList();

                        writer.WriteLine(string.Join(",", visibleColumns.Select(c => EscapeClaimCsvValue(c.HeaderText)).ToArray()));
                        foreach (DataGridViewRow row in _dgvClaimComparison.Rows)
                        {
                            if (row.IsNewRow) continue;
                            writer.WriteLine(string.Join(",", visibleColumns.Select(c => EscapeClaimCsvValue(Convert.ToString(row.Cells[c.Index].Value))).ToArray()));
                        }
                    }

                    ShowToast("청구 비교 결과를 CSV로 저장했습니다.", ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("CSV 저장 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string EscapeClaimCsvValue(string value)
        {
            string text = value ?? "";
            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }






        // =========================================================================================
        // UI & Logic - 로그/청구 기반 검사 및 분리 복구 도구 (Log & Claim Mismatch Scanner)
        // =========================================================================================
        private DataGridView CreateDarkDataGrid()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = ColorBgCard;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = ColorTextMain;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontBold;
            dgv.DefaultCellStyle.BackColor = ColorBgCard;
            dgv.DefaultCellStyle.ForeColor = ColorTextMain;
            dgv.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            return dgv;
        }

        private void LayoutLogMismatchActionControls()
        {
            if (_pnlLogMismatchDetailAction == null ||
                _btnLogRestoreSelectAll == null || _btnLogRestoreDeselectAll == null ||
                _lblLogRestorePatient == null || _cmbLogRestorePatientGroup == null ||
                _lblLogRestoreNewChart == null || _txtLogRestoreNewChrtNo == null ||
                _btnLogRestoreSeparate == null)
            {
                return;
            }

            bool juminMode = _cmbLogMismatchFilter != null && _cmbLogMismatchFilter.SelectedIndex == 7;
            _lblLogRestorePatient.Visible = !juminMode;
            _cmbLogRestorePatientGroup.Visible = !juminMode;
            _lblLogRestoreNewChart.Visible = !juminMode;
            _txtLogRestoreNewChrtNo.Visible = !juminMode;

            if (juminMode)
            {
                int batchX = 6;
                _btnLogRestoreSelectAll.Left = batchX;
                _btnLogRestoreSelectAll.Width = Math.Max(190,
                    TextRenderer.MeasureText(_btnLogRestoreSelectAll.Text, _btnLogRestoreSelectAll.Font,
                        new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width + 50);
                batchX = _btnLogRestoreSelectAll.Right + 6;
                _btnLogRestoreDeselectAll.Left = batchX;
                _btnLogRestoreDeselectAll.Width = Math.Max(95,
                    TextRenderer.MeasureText(_btnLogRestoreDeselectAll.Text, _btnLogRestoreDeselectAll.Font,
                        new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width + 40);
                batchX = _btnLogRestoreDeselectAll.Right + 12;
                _btnLogRestoreSeparate.Left = batchX;
                _btnLogRestoreSeparate.Width = Math.Max(330,
                    TextRenderer.MeasureText(_btnLogRestoreSeparate.Text, _btnLogRestoreSeparate.Font,
                        new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width + 54);
                return;
            }

            // ApplyModernStyleRecursive expands short buttons to fit their text. Reflow the
            // following controls after that expansion so "복구환자" is never covered by "해제".
            int x = 6;
            _btnLogRestoreSelectAll.Left = x;
            x = _btnLogRestoreSelectAll.Right + 6;

            _btnLogRestoreDeselectAll.Left = x;
            x = _btnLogRestoreDeselectAll.Right + 12;

            _lblLogRestorePatient.Left = x;
            _lblLogRestorePatient.Width = Math.Max(72,
                TextRenderer.MeasureText(_lblLogRestorePatient.Text, _lblLogRestorePatient.Font).Width + 6);
            x = _lblLogRestorePatient.Right + 6;

            _cmbLogRestorePatientGroup.Left = x;
            _cmbLogRestorePatientGroup.DropDownWidth = Math.Max(320, _cmbLogRestorePatientGroup.Width);
            x = _cmbLogRestorePatientGroup.Right + 10;

            _lblLogRestoreNewChart.Left = x;
            _lblLogRestoreNewChart.Width = Math.Max(60,
                TextRenderer.MeasureText(_lblLogRestoreNewChart.Text, _lblLogRestoreNewChart.Font).Width + 6);
            x = _lblLogRestoreNewChart.Right + 6;

            _txtLogRestoreNewChrtNo.Left = x;
            x = _txtLogRestoreNewChrtNo.Right + 10;
            _btnLogRestoreSeparate.Left = x;
        }

        private Button CreateJuminViewButton(string text, int width)
        {
            Button button = new Button
            {
                Text = text,
                Size = new Size(width, 27),
                Margin = new Padding(0, 0, 4, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void InitializeLogClaimMismatchTab()
        {
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            _tabLogClaimMismatch.Controls.Add(layout);

            // 1. Top Search & Condition Panel
            Panel pnlSearch = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            layout.Controls.Add(pnlSearch, 0, 0);

            Label lblFilter = new Label
            {
                Text = "검사 유형",
                Location = new Point(14, 14),
                Size = new Size(65, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlSearch.Controls.Add(lblFilter);

            _cmbLogMismatchFilter = new ComboBox
            {
                Location = new Point(80, 10),
                Size = new Size(200, 27),
                DropDownWidth = 340,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            _cmbLogMismatchFilter.Items.AddRange(new object[] {
                "전체 이상 감지 (로그/청구 불일치)",
                "처방로그 환자명 불일치 (원장 vs 로그)",
                "청구/로그 환자 불일치",
                "다중 환자 병합 의심 (1개 차트에 여러 환자)",
                "특정 환자/차트번호 직접 검색",
                "로그없는 처방 포함 (자동판정 불가)",
                "백업 원환자와 불일치",
                "주민번호 암호문 불일치"
            });
            _cmbLogMismatchFilter.SelectedIndex = 0;
            pnlSearch.Controls.Add(_cmbLogMismatchFilter);

            Label lblTarget = new Label
            {
                Text = "검색 대상",
                Location = new Point(295, 14),
                Size = new Size(65, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlSearch.Controls.Add(lblTarget);

            _txtLogMismatchTarget = new TextBox
            {
                Location = new Point(362, 10),
                Size = new Size(160, 27),
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            pnlSearch.Controls.Add(_txtLogMismatchTarget);

            _btnLogMismatchScan = new Button
            {
                Text = "🔍 로그/청구 기반 무결성 검사 실행",
                Location = new Point(535, 8),
                Size = new Size(245, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnLogMismatchScan.FlatAppearance.BorderSize = 0;
            _btnLogMismatchScan.Click += BtnLogMismatchScan_Click;
            pnlSearch.Controls.Add(_btnLogMismatchScan);

            _btnLogMismatchExport = new Button
            {
                Text = "📊 결과 CSV 저장",
                Location = new Point(790, 8),
                Size = new Size(130, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Enabled = false
            };
            _btnLogMismatchExport.FlatAppearance.BorderSize = 0;
            _btnLogMismatchExport.Click += BtnLogMismatchExport_Click;
            pnlSearch.Controls.Add(_btnLogMismatchExport);

            _btnAttachPrescriptionBackup = new Button
            {
                Text = "💾 백업 DB 연결",
                Location = new Point(930, 8),
                Size = new Size(130, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(71, 85, 105),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnAttachPrescriptionBackup.FlatAppearance.BorderSize = 0;
            _btnAttachPrescriptionBackup.Click += BtnAttachPrescriptionBackup_Click;
            pnlSearch.Controls.Add(_btnAttachPrescriptionBackup);

            _btnDetachPrescriptionBackup = new Button
            {
                Text = "❌ 백업 DB 연결 해제",
                Location = new Point(1068, 8),
                Size = new Size(155, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Enabled = false
            };
            _btnDetachPrescriptionBackup.FlatAppearance.BorderSize = 0;
            _btnDetachPrescriptionBackup.Click += BtnDetachPrescriptionBackup_Click;
            pnlSearch.Controls.Add(_btnDetachPrescriptionBackup);

            _lblBackupConnectionStatus = new Label
            {
                Text = "백업 DB 연결 상태를 확인하는 중입니다...",
                Location = new Point(14, 45),
                Size = new Size(1150, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Italic)
            };
            pnlSearch.Controls.Add(_lblBackupConnectionStatus);
            _tabLogClaimMismatch.Enter += (s, e) => RefreshAttachedBackupStatus();

            // 2. Middle SplitContainer (Summary Grid on Left, Detail & Restore on Right)
            _splitLogMismatch = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Size = new Size(1100, 600),
                Panel1MinSize = 100,
                Panel2MinSize = 100,
                BackColor = ColorBorder
            };
            try
            {
                _splitLogMismatch.SplitterDistance = Math.Max(100, Math.Min(600, _distLogMismatch));
            }
            catch { }
            _splitLogMismatch.SplitterMoved += (s, e) => _distLogMismatch = _splitLogMismatch.SplitterDistance;
            layout.Controls.Add(_splitLogMismatch, 0, 1);

            // Panel1: Summary Grid
            Panel pnlSummaryHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(8)
            };
            _splitLogMismatch.Panel1.Controls.Add(pnlSummaryHost);

            TableLayoutPanel summaryLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = new Padding(0),
                Padding = new Padding(0),
                BackColor = ColorBgCard
            };
            summaryLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            summaryLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));
            summaryLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            summaryLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            pnlSummaryHost.Controls.Add(summaryLayout);

            Label lblSummaryHeader = new Label
            {
                Text = "📋 이상 감지 환자/차트 목록 (선택 시 상세 대조)",
                Dock = DockStyle.Fill,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            summaryLayout.Controls.Add(lblSummaryHeader, 0, 0);

            _pnlJuminClassificationViews = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0, 2, 0, 2),
                BackColor = ColorBgMain,
                Visible = false
            };
            _pnlJuminClassificationViews.Resize += (s, e) => ResizeJuminClassificationViewControls();
            summaryLayout.Controls.Add(_pnlJuminClassificationViews, 0, 1);

            _btnJuminShowRestoreTargets = CreateJuminViewButton("복구 대상 0명", 112);
            _btnJuminShowRestoreTargets.Click += (s, e) => ShowJuminClassificationView("복구 대상");
            _pnlJuminClassificationViews.Controls.Add(_btnJuminShowRestoreTargets);

            _btnJuminShowNoEvidence = CreateJuminViewButton("근거 없음 0명", 112);
            _btnJuminShowNoEvidence.Click += (s, e) => ShowJuminClassificationView("근거 없음");
            _pnlJuminClassificationViews.Controls.Add(_btnJuminShowNoEvidence);

            _btnJuminShowUnidentified = CreateJuminViewButton("식별 불가 0명", 112);
            _btnJuminShowUnidentified.Click += (s, e) => ShowJuminClassificationView("식별 불가");
            _pnlJuminClassificationViews.Controls.Add(_btnJuminShowUnidentified);

            _lblJuminNormalCount = new Label
            {
                Text = "이미 정상 0명",
                AutoSize = false,
                Size = new Size(112, 27),
                Margin = new Padding(4, 0, 0, 0),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(23, 37, 56),
                ForeColor = Color.FromArgb(147, 197, 253),
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _pnlJuminClassificationViews.Controls.Add(_lblJuminNormalCount);

            _dgvLogMismatchSummary = CreateDarkDataGrid();
            _dgvLogMismatchSummary.Dock = DockStyle.Fill;
            _dgvLogMismatchSummary.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dgvLogMismatchSummary.MultiSelect = false;
            _dgvLogMismatchSummary.ReadOnly = true;
            _dgvLogMismatchSummary.CellClick += DgvLogMismatchSummary_CellClick;
            _dgvLogMismatchSummary.SelectionChanged += DgvLogMismatchSummary_SelectionChanged;
            _dgvLogMismatchSummary.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_dgvLogMismatchSummary.IsCurrentCellDirty && _dgvLogMismatchSummary.CurrentCell is DataGridViewCheckBoxCell)
                    _dgvLogMismatchSummary.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _dgvLogMismatchSummary.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                    _dgvLogMismatchSummary.Columns[e.ColumnIndex].Name == "암호복구선택")
                    UpdateJuminRestoreButtonState();
            };
            summaryLayout.Controls.Add(_dgvLogMismatchSummary, 0, 2);

            // Panel2: Detail Grid & Separation Action Toolbar
            Panel pnlDetailHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(8)
            };
            _splitLogMismatch.Panel2.Controls.Add(pnlDetailHost);

            // Detail Top Action Bar
            _pnlLogMismatchDetailAction = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                BackColor = ColorBgMain,
                Padding = new Padding(6)
            };
            _pnlLogMismatchDetailAction.Resize += (s, e) => LayoutLogMismatchActionControls();
            pnlDetailHost.Controls.Add(_pnlLogMismatchDetailAction);

            _lblLogMismatchDetailInfo = new Label
            {
                Text = "선택된 차트 없음 (좌측 목록에서 검사 대상 차트를 클릭하십시오)",
                Location = new Point(6, 6),
                Size = new Size(760, 22),
                AutoEllipsis = true,
                ForeColor = ColorEmerald,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold)
            };
            _pnlLogMismatchDetailAction.Controls.Add(_lblLogMismatchDetailInfo);

            _btnLogRestoreSelectAll = new Button
            {
                Text = "✓ 환자별 선택",
                Location = new Point(6, 36),
                Size = new Size(110, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            _btnLogRestoreSelectAll.FlatAppearance.BorderColor = ColorBorder;
            _btnLogRestoreSelectAll.Click += (s, e) =>
            {
                if (_cmbLogMismatchFilter != null && _cmbLogMismatchFilter.SelectedIndex == 7)
                    SetJuminDetailChecked(true);
                else
                    SelectDetailPatientGroup();
            };
            _pnlLogMismatchDetailAction.Controls.Add(_btnLogRestoreSelectAll);

            _btnLogRestoreDeselectAll = new Button
            {
                Text = "✗ 해제",
                Location = new Point(122, 36),
                Size = new Size(62, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            _btnLogRestoreDeselectAll.FlatAppearance.BorderColor = ColorBorder;
            _btnLogRestoreDeselectAll.Click += (s, e) =>
            {
                if (_cmbLogMismatchFilter != null && _cmbLogMismatchFilter.SelectedIndex == 7)
                    SetJuminDetailChecked(false);
                else
                    SetDetailGridChecked(false);
            };
            _pnlLogMismatchDetailAction.Controls.Add(_btnLogRestoreDeselectAll);

            _lblLogRestorePatient = new Label
            {
                Text = "복구환자:",
                Location = new Point(196, 40),
                Size = new Size(72, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _pnlLogMismatchDetailAction.Controls.Add(_lblLogRestorePatient);

            _cmbLogRestorePatientGroup = new ComboBox
            {
                Location = new Point(270, 37),
                Size = new Size(249, 27),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };
            _cmbLogRestorePatientGroup.SelectedIndexChanged += CmbLogRestorePatientGroup_SelectedIndexChanged;
            _pnlLogMismatchDetailAction.Controls.Add(_cmbLogRestorePatientGroup);

            _lblLogRestoreNewChart = new Label
            {
                Text = "원차트:",
                Location = new Point(529, 40),
                Size = new Size(60, 22),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _pnlLogMismatchDetailAction.Controls.Add(_lblLogRestoreNewChart);

            _txtLogRestoreNewChrtNo = new TextBox
            {
                Location = new Point(591, 37),
                Size = new Size(108, 25),
                BackColor = ColorBgCard,
                ForeColor = ColorTextMain,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _pnlLogMismatchDetailAction.Controls.Add(_txtLogRestoreNewChrtNo);

            _btnLogRestoreSeparate = new Button
            {
                Text = "🛠️ 선택 처방을 원차트로 분리/복구",
                Location = new Point(709, 34),
                Size = new Size(300, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnLogRestoreSeparate.FlatAppearance.BorderSize = 0;
            _btnLogRestoreSeparate.Click += BtnLogRestoreSeparate_Click;
            _pnlLogMismatchDetailAction.Controls.Add(_btnLogRestoreSeparate);

            LayoutLogMismatchActionControls();

            // Detail Grid
            _dgvLogMismatchDetail = CreateDarkDataGrid();
            _dgvLogMismatchDetail.Dock = DockStyle.Fill;
            _dgvLogMismatchDetail.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dgvLogMismatchDetail.MultiSelect = false;
            _dgvLogMismatchDetail.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_dgvLogMismatchDetail.IsCurrentCellDirty)
                {
                    _dgvLogMismatchDetail.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
            _dgvLogMismatchDetail.CellValueChanged += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                    _dgvLogMismatchDetail.Columns[e.ColumnIndex].Name == "선택" &&
                    _cmbLogMismatchFilter != null && _cmbLogMismatchFilter.SelectedIndex == 7)
                {
                    UpdateJuminRestoreButtonState();
                }
            };
            _dgvLogMismatchDetail.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0 || _dgvLogMismatchDetail.Columns["복구원차트"] == null) return;
                string originalChart = Convert.ToString(_dgvLogMismatchDetail.Rows[e.RowIndex].Cells["복구원차트"].Value).Trim();
                string patientName = Convert.ToString(_dgvLogMismatchDetail.Rows[e.RowIndex].Cells["복구환자명"].Value).Trim();
                if (!string.IsNullOrEmpty(originalChart)) _txtLogRestoreNewChrtNo.Text = originalChart;
                SelectRestorePatientGroupInCombo(patientName, originalChart);
            };
            pnlDetailHost.Controls.Add(_dgvLogMismatchDetail);
            _dgvLogMismatchDetail.BringToFront();

            // 3. Bottom Status Label
            _lblLogMismatchSummary = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = ColorBgCard,
                ForeColor = ColorTextSec,
                Padding = new Padding(12, 0, 0, 0),
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular),
                Text = "로그/청구 검사 대기 중"
            };
            layout.Controls.Add(_lblLogMismatchSummary, 0, 2);
        }

        private void PopulateRestorePatientGroups(DataTable detailDt)
        {
            if (_cmbLogRestorePatientGroup == null) return;

            DataTable groups = new DataTable();
            groups.Columns.Add("환자명");
            groups.Columns.Add("원차트");
            groups.Columns.Add("건수", typeof(int));
            groups.Columns.Add("표시");

            Dictionary<string, DataRow> rowsByKey = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in detailDt.Rows)
            {
                if (!Convert.ToBoolean(row["복구가능"])) continue;
                string patientName = Convert.ToString(row["복구환자명"]).Trim();
                string chartNo = Convert.ToString(row["복구원차트"]).Trim();
                if (string.IsNullOrEmpty(patientName) || string.IsNullOrEmpty(chartNo)) continue;

                string key = patientName + "\u001f" + chartNo;
                DataRow groupRow;
                if (!rowsByKey.TryGetValue(key, out groupRow))
                {
                    groupRow = groups.NewRow();
                    groupRow["환자명"] = patientName;
                    groupRow["원차트"] = chartNo;
                    groupRow["건수"] = 0;
                    groups.Rows.Add(groupRow);
                    rowsByKey[key] = groupRow;
                }
                groupRow["건수"] = Convert.ToInt32(groupRow["건수"]) + 1;
            }

            foreach (DataRow row in groups.Rows)
            {
                row["표시"] = string.Format("{0} / {1} ({2}건)", row["환자명"], row["원차트"], row["건수"]);
            }

            _cmbLogRestorePatientGroup.DataSource = null;
            _cmbLogRestorePatientGroup.DisplayMember = "표시";
            _cmbLogRestorePatientGroup.DataSource = groups;
            _cmbLogRestorePatientGroup.SelectedIndex = groups.Rows.Count == 1 ? 0 : -1;
        }

        private void CmbLogRestorePatientGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataRowView selected = _cmbLogRestorePatientGroup == null ? null : _cmbLogRestorePatientGroup.SelectedItem as DataRowView;
            if (selected == null) return;
            string chartNo = Convert.ToString(selected["원차트"]).Trim();
            if (!string.IsNullOrEmpty(chartNo)) _txtLogRestoreNewChrtNo.Text = chartNo;
        }

        private void SelectRestorePatientGroupInCombo(string patientName, string chartNo)
        {
            if (_cmbLogRestorePatientGroup == null || _cmbLogRestorePatientGroup.DataSource == null) return;
            for (int i = 0; i < _cmbLogRestorePatientGroup.Items.Count; i++)
            {
                DataRowView item = _cmbLogRestorePatientGroup.Items[i] as DataRowView;
                if (item == null) continue;
                if (string.Equals(Convert.ToString(item["환자명"]).Trim(), patientName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Convert.ToString(item["원차트"]).Trim(), chartNo, StringComparison.OrdinalIgnoreCase))
                {
                    _cmbLogRestorePatientGroup.SelectedIndex = i;
                    return;
                }
            }
        }

        private void SelectDetailPatientGroup()
        {
            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Rows.Count == 0) return;
            DataRowView selected = _cmbLogRestorePatientGroup == null ? null : _cmbLogRestorePatientGroup.SelectedItem as DataRowView;
            if (selected == null)
            {
                MessageBox.Show("복구환자 목록에서 환자와 원차트를 먼저 선택하십시오.", "복구환자 선택", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string patientName = Convert.ToString(selected["환자명"]).Trim();
            string chartNo = Convert.ToString(selected["원차트"]).Trim();
            int selectedCount = 0;
            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                bool restorable = row.Cells["복구가능"] != null && Convert.ToBoolean(row.Cells["복구가능"].Value);
                bool samePatient = string.Equals(Convert.ToString(row.Cells["복구환자명"].Value).Trim(), patientName, StringComparison.OrdinalIgnoreCase);
                bool sameChart = string.Equals(Convert.ToString(row.Cells["복구원차트"].Value).Trim(), chartNo, StringComparison.OrdinalIgnoreCase);
                bool shouldSelect = restorable && samePatient && sameChart;
                row.Cells["선택"].Value = shouldSelect;
                if (shouldSelect) selectedCount++;
            }

            _txtLogRestoreNewChrtNo.Text = chartNo;
            ShowToast(string.Format("{0} / {1}: {2}건을 선택했습니다.", patientName, chartNo, selectedCount), ColorEmerald);
        }

        private void SetDetailGridChecked(bool check)
        {
            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Rows.Count == 0) return;
            string targetOriginalChart = _txtLogRestoreNewChrtNo == null ? "" : _txtLogRestoreNewChrtNo.Text.Trim();
            if (check && string.IsNullOrEmpty(targetOriginalChart) && _dgvLogMismatchDetail.CurrentRow != null)
            {
                targetOriginalChart = Convert.ToString(_dgvLogMismatchDetail.CurrentRow.Cells["복구원차트"].Value).Trim();
                if (!string.IsNullOrEmpty(targetOriginalChart)) _txtLogRestoreNewChrtNo.Text = targetOriginalChart;
            }
            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                if (row.Cells["선택"] != null)
                {
                    bool restorable = row.Cells["복구가능"] != null && Convert.ToBoolean(row.Cells["복구가능"].Value);
                    string rowOriginalChart = Convert.ToString(row.Cells["복구원차트"].Value).Trim();
                    bool sameOriginalChart = string.IsNullOrEmpty(targetOriginalChart) ||
                                             string.Equals(targetOriginalChart, rowOriginalChart, StringComparison.OrdinalIgnoreCase);
                    row.Cells["선택"].Value = check && restorable && sameOriginalChart;
                }
            }
        }

        private void ConfigureLogMismatchDetailHeaders(bool juminEncryptionMode)
        {
            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Columns.Count == 0) return;

            Dictionary<string, string> headers = juminEncryptionMode
                ? new Dictionary<string, string>
                {
                    { "선택", "암호복구선택" },
                    { "복구가능", "암호복구가능" },
                    { "원장차트", "현재차트(변경안함)" },
                    { "복구근거", "차트근거(참고)" },
                    { "복구원차트", "과거차트(참고)" },
                    { "복구환자명", "근거환자명(참고)" },
                    { "복구주민번호", "근거주민번호(참고)" },
                    { "진단상태", "차트판정(참고)" }
                }
                : new Dictionary<string, string>
                {
                    { "선택", "선택" },
                    { "복구가능", "복구가능" },
                    { "원장차트", "원장차트" },
                    { "복구근거", "복구근거" },
                    { "복구원차트", "복구원차트" },
                    { "복구환자명", "복구환자명" },
                    { "복구주민번호", "복구주민번호" },
                    { "진단상태", "진단상태" }
                };

            foreach (KeyValuePair<string, string> item in headers)
            {
                if (_dgvLogMismatchDetail.Columns.Contains(item.Key))
                    _dgvLogMismatchDetail.Columns[item.Key].HeaderText = item.Value;
            }

            if (juminEncryptionMode)
            {
                if (_dgvLogMismatchDetail.Columns.Contains("선택"))
                    _dgvLogMismatchDetail.Columns["선택"].ToolTipText = "오른쪽 처방 전체를 선택해야 환자 단위 암호문 복구가 가능합니다.";
                if (_dgvLogMismatchDetail.Columns.Contains("복구근거"))
                    _dgvLogMismatchDetail.Columns["복구근거"].ToolTipText = "차트 소유권 판정용 참고정보입니다. 암호문 복구 근거와는 별개입니다.";
                if (_dgvLogMismatchDetail.Columns.Contains("복구원차트"))
                    _dgvLogMismatchDetail.Columns["복구원차트"].ToolTipText = "표시된 과거 차트로 이동하지 않으며 현재 차트번호는 변경하지 않습니다.";
            }
        }

        private void BtnAttachPrescriptionBackup_Click(object sender, EventArgs e)
        {
            if (_chkDemoMode.Checked)
            {
                MessageBox.Show("실서버 모드에서만 백업 DB를 연결할 수 있습니다.", "백업 DB 연결", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "PM_MAIN.MDF와 PM_MAIN_LOG.LDF가 있는 백업 폴더를 선택하십시오.";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    _lblLogMismatchSummary.Text = "백업 파일을 검사하는 중입니다...";
                    Application.DoEvents();

                    string sourceMdf = Directory.GetFiles(dialog.SelectedPath)
                        .FirstOrDefault(p => string.Equals(Path.GetFileName(p), "PM_MAIN.MDF", StringComparison.OrdinalIgnoreCase));
                    string sourceLdf = Directory.GetFiles(dialog.SelectedPath)
                        .FirstOrDefault(p => Path.GetExtension(p).Equals(".ldf", StringComparison.OrdinalIgnoreCase)
                            && Path.GetFileNameWithoutExtension(p).StartsWith("PM_MAIN", StringComparison.OrdinalIgnoreCase));
                    if (string.IsNullOrEmpty(sourceMdf) || string.IsNullOrEmpty(sourceLdf))
                    {
                        throw new FileNotFoundException("선택한 폴더에서 PM_MAIN.MDF와 PM_MAIN 계열 LDF 파일을 모두 찾지 못했습니다.");
                    }

                    Match dateMatch = Regex.Match(Path.GetFileName(dialog.SelectedPath.TrimEnd(Path.DirectorySeparatorChar)), @"20\d{6}");
                    string backupDate = dateMatch.Success ? dateMatch.Value : File.GetLastWriteTime(sourceMdf).ToString("yyyyMMdd");
                    string databaseName = "PM_MAIN_BACKUP_" + backupDate;

                    SqlConnectionStringBuilder masterBuilder = new SqlConnectionStringBuilder(BuildConnectionString(false));
                    masterBuilder.InitialCatalog = "master";
                    using (SqlConnection conn = new SqlConnection(masterBuilder.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand existsCmd = new SqlCommand("SELECT COUNT(*) FROM sys.databases WHERE name = @name;", conn))
                        {
                            existsCmd.Parameters.Add("@name", SqlDbType.NVarChar, 128).Value = databaseName;
                            if (Convert.ToInt32(existsCmd.ExecuteScalar()) > 0)
                            {
                                MessageBox.Show(string.Format("백업 DB [{0}]가 이미 연결되어 있습니다. 바로 검사를 실행할 수 있습니다.", databaseName), "백업 DB 연결", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                _lblLogMismatchSummary.Text = "연결된 백업 DB: " + databaseName + " (읽기 전용)";
                                RefreshAttachedBackupStatus();
                                return;
                            }
                        }

                        string dataPath;
                        string logPath;
                        using (SqlCommand pathCmd = new SqlCommand(@"
SELECT CAST(SERVERPROPERTY('InstanceDefaultDataPath') AS NVARCHAR(4000)),
       CAST(SERVERPROPERTY('InstanceDefaultLogPath') AS NVARCHAR(4000));", conn))
                        using (SqlDataReader reader = pathCmd.ExecuteReader())
                        {
                            if (!reader.Read()) throw new InvalidOperationException("SQL Server 기본 데이터 경로를 확인하지 못했습니다.");
                            dataPath = Convert.ToString(reader[0]);
                            logPath = Convert.ToString(reader[1]);
                        }

                        if (string.IsNullOrEmpty(dataPath)) throw new InvalidOperationException("SQL Server 기본 데이터 경로가 비어 있습니다.");
                        if (string.IsNullOrEmpty(logPath)) logPath = dataPath;
                        string copiedMdf = Path.Combine(dataPath, databaseName + ".mdf");
                        string copiedLdf = Path.Combine(logPath, databaseName + "_log.ldf");

                        if (File.Exists(copiedMdf) || File.Exists(copiedLdf))
                        {
                            throw new IOException("같은 이름의 복사 파일이 SQL 데이터 폴더에 이미 있습니다. DB 연결 상태를 확인한 뒤 처리하십시오: " + databaseName);
                        }

                        long requiredBytes = new FileInfo(sourceMdf).Length + new FileInfo(sourceLdf).Length + (200L * 1024L * 1024L);
                        DriveInfo drive = new DriveInfo(Path.GetPathRoot(dataPath));
                        if (drive.AvailableFreeSpace < requiredBytes)
                        {
                            throw new IOException(string.Format("SQL 데이터 드라이브의 여유 공간이 부족합니다. 최소 {0:N1}GB가 필요합니다.", requiredBytes / 1073741824D));
                        }

                        _lblLogMismatchSummary.Text = string.Format("백업 원본을 보존하기 위해 {0:N1}GB 복사 중...", (new FileInfo(sourceMdf).Length + new FileInfo(sourceLdf).Length) / 1073741824D);
                        Application.DoEvents();
                        File.Copy(sourceMdf, copiedMdf, false);
                        File.Copy(sourceLdf, copiedLdf, false);

                        string attachSql = "CREATE DATABASE " + QuoteSqlName(databaseName)
                            + " ON (FILENAME=N'" + copiedMdf.Replace("'", "''") + "'),"
                            + " (FILENAME=N'" + copiedLdf.Replace("'", "''") + "') FOR ATTACH;"
                            + " ALTER DATABASE " + QuoteSqlName(databaseName) + " SET READ_ONLY WITH ROLLBACK IMMEDIATE;";
                        using (SqlCommand attachCmd = new SqlCommand(attachSql, conn))
                        {
                            attachCmd.CommandTimeout = 180;
                            attachCmd.ExecuteNonQuery();
                        }
                    }

                    _lblLogMismatchSummary.Text = "연결된 백업 DB: " + databaseName + " (읽기 전용)";
                    RefreshAttachedBackupStatus();
                    MessageBox.Show(
                        string.Format("백업 DB 연결이 완료되었습니다.\n\n- DB 이름: {0}\n- 원본 백업: 변경하지 않음\n- 연결 상태: 읽기 전용\n\n이제 무결성 검사를 실행하십시오.", databaseName),
                        "백업 DB 연결 완료",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (UnauthorizedAccessException ex)
                {
                    MessageBox.Show("SQL 데이터 폴더에 복사할 권한이 없습니다. 프로그램을 관리자 권한으로 실행한 뒤 다시 시도하십시오.\n\n" + ex.Message, "관리자 권한 필요", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("백업 DB 연결 중 오류가 발생했습니다. 운영 DB는 변경하지 않았습니다.\n\n" + ex.Message, "백업 DB 연결 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void RefreshAttachedBackupStatus()
        {
            if (_lblBackupConnectionStatus == null) return;

            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                _lblBackupConnectionStatus.Text = "● 데모 모드: 백업 DB를 사용하지 않습니다.";
                _lblBackupConnectionStatus.ForeColor = ColorWarning;
                if (_btnDetachPrescriptionBackup != null) _btnDetachPrescriptionBackup.Enabled = false;
                return;
            }

            bool readOnly;
            string databaseName = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (_btnDetachPrescriptionBackup != null)
            {
                _btnDetachPrescriptionBackup.Enabled = !string.IsNullOrEmpty(databaseName);
            }

            if (string.IsNullOrEmpty(databaseName))
            {
                _lblBackupConnectionStatus.Text = "● 백업 DB 미연결: [백업 DB 연결]로 PM_MAIN.MDF 폴더를 선택하십시오.";
                _lblBackupConnectionStatus.ForeColor = ColorWarning;
            }
            else
            {
                _lblBackupConnectionStatus.Text = string.Format(
                    "● 연결된 백업 DB: {0}  |  상태: {1}",
                    databaseName,
                    readOnly ? "읽기 전용" : "주의 - 쓰기 가능");
                _lblBackupConnectionStatus.ForeColor = readOnly ? ColorEmerald : ColorAlarm;
            }
        }

        private void BtnDetachPrescriptionBackup_Click(object sender, EventArgs e)
        {
            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                MessageBox.Show("데모 모드에서는 백업 DB 연결 해제 기능을 사용하지 않습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool readOnly;
            string databaseName = FindAttachedPrescriptionBackupDatabase(out readOnly);
            if (string.IsNullOrEmpty(databaseName))
            {
                MessageBox.Show("현재 연결된 백업 데이터베이스가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                RefreshAttachedBackupStatus();
                return;
            }

            DialogResult dr = MessageBox.Show(
                string.Format(
                    "현재 연결된 백업 데이터베이스 [{0}]의 연결을 해제(분리)하시겠습니까?\n\n" +
                    "※ 분리(Detach) 시 SQL Server 연결이 정상 해제되며, 임시 복사 파일도 안전하게 정리됩니다.\n" +
                    "(D:\\Downloads 등의 원본 백업 파일은 절대 삭제되지 않고 온전히 유지됩니다.)",
                    databaseName),
                "백업 DB 연결 해제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dr != DialogResult.Yes) return;

            try
            {
                this.Cursor = Cursors.WaitCursor;
                _lblBackupConnectionStatus.Text = string.Format("백업 DB [{0}] 연결 해제 중...", databaseName);
                Application.DoEvents();

                string mdfPath = "";
                string ldfPath = "";

                SqlConnectionStringBuilder masterBuilder = new SqlConnectionStringBuilder(BuildConnectionString(false));
                masterBuilder.InitialCatalog = "master";

                using (SqlConnection conn = new SqlConnection(masterBuilder.ConnectionString))
                {
                    conn.Open();

                    // 1. 물리 파일 경로 확인 (sys.master_files)
                    try
                    {
                        using (SqlCommand pathCmd = new SqlCommand(@"
                            SELECT physical_name 
                            FROM sys.master_files 
                            WHERE database_id = DB_ID(@dbname);", conn))
                        {
                            pathCmd.Parameters.Add("@dbname", SqlDbType.NVarChar, 128).Value = databaseName;
                            using (SqlDataReader reader = pathCmd.ExecuteReader())
                            {
                                if (reader.Read()) mdfPath = reader.IsDBNull(0) ? "" : reader.GetString(0);
                                if (reader.Read()) ldfPath = reader.IsDBNull(0) ? "" : reader.GetString(0);
                            }
                        }
                    }
                    catch { }

                    // 2. 단일 사용자 모드로 변경 후 분리(sp_detach_db)
                    string detachSql = string.Format(@"
                        ALTER DATABASE {0} SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
                        EXEC master.dbo.sp_detach_db @dbname = N'{1}', @skipchecks = 'true';",
                        QuoteSqlName(databaseName), databaseName.Replace("'", "''"));

                    using (SqlCommand detachCmd = new SqlCommand(detachSql, conn))
                    {
                        detachCmd.CommandTimeout = 60;
                        detachCmd.ExecuteNonQuery();
                    }
                }

                // 3. SQL 데이터 폴더에 임시 복사되었던 파일 안전 삭제
                try
                {
                    if (!string.IsNullOrEmpty(mdfPath) && File.Exists(mdfPath) && mdfPath.IndexOf("PM_MAIN_BACKUP_", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        File.Delete(mdfPath);
                    }
                }
                catch { }

                try
                {
                    if (!string.IsNullOrEmpty(ldfPath) && File.Exists(ldfPath) && ldfPath.IndexOf("PM_MAIN_BACKUP_", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        File.Delete(ldfPath);
                    }
                }
                catch { }

                RefreshAttachedBackupStatus();
                if (_lblLogMismatchSummary != null)
                {
                    _lblLogMismatchSummary.Text = "백업 DB 연결이 정상적으로 해제되었습니다.";
                }

                ShowToast("백업 DB 연결 해제 완료", ColorEmerald);
                MessageBox.Show(
                    string.Format("백업 데이터베이스 [{0}]의 연결이 정상적으로 해제(분리)되었습니다.", databaseName),
                    "백업 DB 연결 해제 완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("백업 DB 연결 해제 중 오류가 발생했습니다:\n\n" + ex.Message, "연결 해제 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RefreshAttachedBackupStatus();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private string FindAttachedPrescriptionBackupDatabase()
        {
            bool readOnly;
            return FindAttachedPrescriptionBackupDatabase(out readOnly);
        }

        private string FindAttachedPrescriptionBackupDatabase(out bool readOnly)
        {
            readOnly = false;
            if (_chkDemoMode != null && _chkDemoMode.Checked) return "";

            try
            {
                using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                using (SqlCommand cmd = new SqlCommand(@"
SELECT TOP (1) d.name, d.is_read_only
FROM sys.databases d
WHERE d.state = 0
  AND d.name LIKE N'PM[_]MAIN[_]BACKUP[_]%'
  AND HAS_DBACCESS(d.name) = 1
ORDER BY d.name DESC;", conn))
                {
                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read()) return "";
                        readOnly = !reader.IsDBNull(1) && reader.GetBoolean(1);
                        return reader.IsDBNull(0) ? "" : reader.GetString(0);
                    }
                }
            }
            catch
            {
                return "";
            }
        }

        private string BuildPrescriptionBackupCte(string databaseName)
        {
            if (string.IsNullOrEmpty(databaseName))
            {
                return @"BackupPres AS
(
    SELECT CAST(NULL AS NVARCHAR(50)) AS DRUG_SEQ,
           CAST(NULL AS NVARCHAR(20)) AS CHRTNO,
           CAST(NULL AS NVARCHAR(100)) AS PAT_NM,
           CAST(NULL AS NVARCHAR(50)) AS PAT_JUMIN_NO
    WHERE 1 = 0
)";
            }

            return @"BackupPres AS
(
    SELECT DRUG_SEQ, CHRTNO, PAT_NM, PAT_JUMIN_NO
    FROM " + QuoteSqlName(databaseName) + @".dbo.TBSID040_03 WITH (NOLOCK)
)";
        }

        private List<string> GetBackupCustomerCloneColumns(SqlConnection conn, SqlTransaction trans, string backupDatabaseName)
        {
            List<string> columns = new List<string>();
            string sql = @"
SELECT c.name
FROM sys.columns c
INNER JOIN sys.tables t ON t.object_id = c.object_id
INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
INNER JOIN " + QuoteSqlName(backupDatabaseName) + @".sys.tables bt ON bt.name = t.name
INNER JOIN " + QuoteSqlName(backupDatabaseName) + @".sys.schemas bs
    ON bs.schema_id = bt.schema_id AND bs.name = s.name
INNER JOIN " + QuoteSqlName(backupDatabaseName) + @".sys.columns bc
    ON bc.object_id = bt.object_id AND bc.name = c.name
WHERE s.name = N'dbo'
  AND t.name = N'TBSIT000_01'
  AND c.is_identity = 0
  AND c.is_computed = 0
  AND bc.is_computed = 0
  AND c.system_type_id <> 189
  AND bc.system_type_id <> 189
ORDER BY c.column_id;";

            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            using (SqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string columnName = reader.IsDBNull(0) ? "" : reader.GetString(0);
                    if (!string.IsNullOrWhiteSpace(columnName)) columns.Add(columnName);
                }
            }

            string[] required = new string[]
            {
                "CHRTNO", "PAT_SEQ", "PAT_NM", "JUMIN_NO", "JUMIN_ENCRYPT", "FAM_NM", "CUSACT"
            };
            foreach (string requiredColumn in required)
            {
                if (!columns.Any(c => string.Equals(c, requiredColumn, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new InvalidOperationException(
                        "운영 DB와 백업 DB의 고객 테이블 공통 컬럼에서 필수 컬럼 [" + requiredColumn + "]을 찾지 못했습니다.");
                }
            }
            return columns;
        }

        private void LoadJuminEncryptionMismatchSummary(DataTable summaryDt, string searchTarget)
        {
            string backupDatabaseName = FindAttachedPrescriptionBackupDatabase();
            if (string.IsNullOrEmpty(backupDatabaseName))
            {
                throw new InvalidOperationException(
                    "주민번호 암호문 검사는 연결된 PM_MAIN 백업 DB가 필요합니다. " +
                    "먼저 [백업 DB 연결]로 백업을 연결한 뒤 다시 실행하십시오.");
            }

            // JUMIN_NO is masked even for healthy patients, so list every active chart and
            // classify it by identity/cipher evidence. Only "복구 가능" rows may be changed.
            string sql = @"
WITH CurrentRaw AS
(
    SELECT CHRTNO, PAT_NM, JUMIN_NO,
           CONVERT(nvarchar(4000), JUMIN_ENCRYPT) AS JUMIN_ENCRYPT, PROC_DTIME,
           LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, ''), '-', ''), ' ', ''), 7) AS JUMIN_PREFIX
    FROM dbo.TBSIT000_01 WITH (NOLOCK)
    WHERE CUSACT = '1'
),
CurrentPatient AS
(
    SELECT CHRTNO, MAX(PAT_NM) AS PAT_NM, MAX(JUMIN_NO) AS JUMIN_NO,
           MAX(JUMIN_ENCRYPT) AS JUMIN_ENCRYPT, MAX(PROC_DTIME) AS PROC_DTIME,
           MAX(JUMIN_PREFIX) AS JUMIN_PREFIX,
           COUNT(DISTINCT LTRIM(RTRIM(ISNULL(PAT_NM, N'')))) AS CURRENT_NAME_COUNT,
           COUNT(DISTINCT JUMIN_PREFIX) AS CURRENT_PREFIX_COUNT,
           COUNT(DISTINCT ISNULL(NULLIF(LTRIM(RTRIM(JUMIN_ENCRYPT)), N''), N'<EMPTY>')) AS CURRENT_CIPHER_COUNT
    FROM CurrentRaw
    GROUP BY CHRTNO
),
BackupPatient AS
(
    SELECT CHRTNO, PAT_NM, JUMIN_NO,
           CONVERT(nvarchar(4000), JUMIN_ENCRYPT) AS JUMIN_ENCRYPT, PROC_DTIME,
           LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, ''), '-', ''), ' ', ''), 7) AS JUMIN_PREFIX
    FROM " + QuoteSqlName(backupDatabaseName) + @".dbo.TBSIT000_01 WITH (NOLOCK)
    WHERE CUSACT = '1'
),
BackupIdentity AS
(
    SELECT LTRIM(RTRIM(ISNULL(PAT_NM, N''))) AS PAT_NM_KEY, JUMIN_PREFIX,
           COUNT(*) AS BACKUP_ROW_COUNT,
           COUNT(DISTINCT CHRTNO) AS BACKUP_CHART_COUNT,
           MIN(CHRTNO) AS BACKUP_FIRST_CHRTNO,
           COUNT(DISTINCT NULLIF(LTRIM(RTRIM(JUMIN_ENCRYPT)), N'')) AS BACKUP_CIPHER_COUNT,
           MAX(NULLIF(LTRIM(RTRIM(JUMIN_ENCRYPT)), N'')) AS BACKUP_JUMIN_ENCRYPT
    FROM BackupPatient
    GROUP BY LTRIM(RTRIM(ISNULL(PAT_NM, N''))), JUMIN_PREFIX
),
UnresolvedBackupMismatch AS
(
    SELECT r.CHRTNO, COUNT(*) AS MISMATCH_COUNT
    FROM dbo.TBSID040_03 r WITH (NOLOCK)
    INNER JOIN " + QuoteSqlName(backupDatabaseName) + @".dbo.TBSID040_03 b WITH (NOLOCK)
        ON b.DRUG_SEQ = r.DRUG_SEQ
    WHERE LTRIM(RTRIM(ISNULL(r.PAT_NM, N''))) <> LTRIM(RTRIM(ISNULL(b.PAT_NM, N'')))
       OR (
            NULLIF(LEFT(REPLACE(REPLACE(ISNULL(r.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7), N'') IS NOT NULL
        AND NULLIF(LEFT(REPLACE(REPLACE(ISNULL(b.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7), N'') IS NOT NULL
        AND LEFT(REPLACE(REPLACE(ISNULL(r.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7)
            <> LEFT(REPLACE(REPLACE(ISNULL(b.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7)
          )
    GROUP BY r.CHRTNO
),
Compared AS
(
    SELECT c.CHRTNO, c.PAT_NM, c.JUMIN_NO, c.JUMIN_ENCRYPT, c.PROC_DTIME, c.JUMIN_PREFIX,
           c.CURRENT_NAME_COUNT, c.CURRENT_PREFIX_COUNT, c.CURRENT_CIPHER_COUNT,
           ISNULL(b.BACKUP_ROW_COUNT, 0) AS BACKUP_ROW_COUNT,
           ISNULL(b.BACKUP_CHART_COUNT, 0) AS BACKUP_CHART_COUNT,
           b.BACKUP_FIRST_CHRTNO,
           ISNULL(b.BACKUP_CIPHER_COUNT, 0) AS BACKUP_CIPHER_COUNT,
           b.BACKUP_JUMIN_ENCRYPT,
           ISNULL(ubm.MISMATCH_COUNT, 0) AS UNRESOLVED_BACKUP_MISMATCH_COUNT,
           CASE WHEN b.BACKUP_CIPHER_COUNT = 1
                  AND ISNULL(c.JUMIN_ENCRYPT, N'') = ISNULL(b.BACKUP_JUMIN_ENCRYPT, N'') THEN 1 ELSE 0 END AS HAS_EXACT_CIPHER
    FROM CurrentPatient c
    LEFT JOIN BackupIdentity b
        ON b.PAT_NM_KEY = LTRIM(RTRIM(c.PAT_NM))
       AND b.JUMIN_PREFIX = c.JUMIN_PREFIX
    LEFT JOIN UnresolvedBackupMismatch ubm ON ubm.CHRTNO = c.CHRTNO
),
RxSummary AS
(
    SELECT CHRTNO, COUNT(*) AS RX_COUNT, MAX(PRES_DTIME) AS LAST_PRES_DTIME
    FROM dbo.TBSID040_03 WITH (NOLOCK)
    GROUP BY CHRTNO
),
Classified AS
(
    SELECT c.*,
           CASE
               WHEN LEN(ISNULL(c.JUMIN_PREFIX, '')) <> 7 OR c.JUMIN_PREFIX LIKE '%[^0-9]%' THEN N'식별 불가'
               WHEN c.CURRENT_NAME_COUNT <> 1 OR c.CURRENT_PREFIX_COUNT <> 1 THEN N'식별 불가'
               WHEN c.UNRESOLVED_BACKUP_MISMATCH_COUNT > 0 THEN N'복구 불가'
               WHEN c.CURRENT_CIPHER_COUNT <> 1 THEN N'복구 불가'
               WHEN c.BACKUP_ROW_COUNT = 0 THEN N'근거 없음'
               WHEN c.BACKUP_CIPHER_COUNT = 0 THEN N'근거 없음'
               WHEN c.BACKUP_CIPHER_COUNT > 1 THEN N'복구 불가'
               WHEN c.HAS_EXACT_CIPHER = 1 THEN N'이미 정상'
               ELSE N'복구 가능'
           END AS RESTORE_STATUS,
           CASE
               WHEN LEN(ISNULL(c.JUMIN_PREFIX, '')) <> 7 OR c.JUMIN_PREFIX LIKE '%[^0-9]%' THEN N'주민번호 앞 7자리 부족 또는 형식 오류'
               WHEN c.CURRENT_NAME_COUNT <> 1 OR c.CURRENT_PREFIX_COUNT <> 1 THEN N'현재 차트에 서로 다른 환자 식별값 존재'
               WHEN c.UNRESOLVED_BACKUP_MISMATCH_COUNT > 0 THEN N'백업 원환자 불일치 ' + CONVERT(nvarchar(20), c.UNRESOLVED_BACKUP_MISMATCH_COUNT) + N'건 먼저 복구 필요'
               WHEN c.CURRENT_CIPHER_COUNT <> 1 THEN N'현재 차트의 암호문이 여러 종류'
               WHEN c.BACKUP_ROW_COUNT = 0 THEN N'백업에서 같은 이름·주민번호 앞 7자리 환자 없음'
               WHEN c.BACKUP_CIPHER_COUNT = 0 THEN N'백업 환자의 암호문 없음'
               WHEN c.BACKUP_CIPHER_COUNT > 1 THEN N'백업 후보 암호문이 여러 종류'
               WHEN c.HAS_EXACT_CIPHER = 1 THEN N'현재 암호문과 백업 암호문 일치'
               ELSE N'백업의 고유 암호문 1개로 복구 가능'
           END AS STATUS_REASON,
           CASE
               WHEN LEN(ISNULL(c.JUMIN_PREFIX, '')) = 7 AND c.JUMIN_PREFIX NOT LIKE '%[^0-9]%'
                AND c.CURRENT_NAME_COUNT = 1 AND c.CURRENT_PREFIX_COUNT = 1 AND c.CURRENT_CIPHER_COUNT = 1
                AND c.UNRESOLVED_BACKUP_MISMATCH_COUNT = 0
                AND c.BACKUP_ROW_COUNT > 0 AND c.BACKUP_CIPHER_COUNT = 1 AND c.HAS_EXACT_CIPHER = 0 THEN 1
               WHEN c.UNRESOLVED_BACKUP_MISMATCH_COUNT > 0 OR c.BACKUP_CIPHER_COUNT > 1 OR c.CURRENT_CIPHER_COUNT <> 1 THEN 2
               WHEN c.BACKUP_ROW_COUNT = 0 OR c.BACKUP_CIPHER_COUNT = 0 THEN 3
               WHEN LEN(ISNULL(c.JUMIN_PREFIX, '')) <> 7 OR c.JUMIN_PREFIX LIKE '%[^0-9]%'
                 OR c.CURRENT_NAME_COUNT <> 1 OR c.CURRENT_PREFIX_COUNT <> 1 THEN 4
               ELSE 5
           END AS STATUS_ORDER
    FROM Compared c
)
SELECT
    s.CHRTNO AS [차트번호], s.PAT_NM AS [현재환자명], s.JUMIN_NO AS [주민등록번호],
    CAST(N'' AS NVARCHAR(100)) AS [로그환자명], CAST(N'' AS NVARCHAR(100)) AS [로그원차트],
    CASE WHEN s.BACKUP_ROW_COUNT > 0 THEN s.PAT_NM ELSE N'' END AS [백업환자명],
    CASE WHEN s.BACKUP_CHART_COUNT = 0 THEN N''
         WHEN s.BACKUP_CHART_COUNT = 1 THEN s.BACKUP_FIRST_CHRTNO
         ELSE s.BACKUP_FIRST_CHRTNO + N' 외 ' + CONVERT(nvarchar(20), s.BACKUP_CHART_COUNT - 1) + N'개'
    END AS [백업원차트],
    s.BACKUP_CHART_COUNT AS [백업확인건수], ISNULL(rx.RX_COUNT, 0) AS [전체처방건수],
    CAST(0 AS INT) AS [로그확인건수], CAST(0 AS INT) AS [로그미확인건수],
    CASE WHEN s.RESTORE_STATUS = N'이미 정상' THEN 0 ELSE 1 END AS [이상건수], CAST(N'' AS NVARCHAR(100)) AS [청구환자명],
    CAST(N'' AS NVARCHAR(100)) AS [청구차트],
    s.STATUS_REASON AS [충돌유형],
    ISNULL(rx.LAST_PRES_DTIME, s.PROC_DTIME) AS [최근조제일],
    s.RESTORE_STATUS AS [복구가능],
    s.UNRESOLVED_BACKUP_MISMATCH_COUNT AS [선행복구필요건수],
    LEN(ISNULL(s.JUMIN_ENCRYPT, N'')) AS [현재암호길이], LEN(ISNULL(s.BACKUP_JUMIN_ENCRYPT, N'')) AS [백업암호길이]
FROM Classified s
LEFT JOIN RxSummary rx ON rx.CHRTNO = s.CHRTNO
WHERE @target = '' OR s.CHRTNO LIKE '%' + @target + '%' OR s.PAT_NM LIKE '%' + @target + '%'
   OR s.JUMIN_PREFIX LIKE @target + '%'
   OR s.BACKUP_FIRST_CHRTNO LIKE '%' + @target + '%'
ORDER BY s.STATUS_ORDER, [최근조제일] DESC, [현재환자명], [차트번호];";

            using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                cmd.CommandTimeout = 300;
                cmd.Parameters.Add("@target", SqlDbType.NVarChar, 50).Value = searchTarget;
                conn.Open();
                adapter.Fill(summaryDt);
            }
        }

        private void BtnLogMismatchScan_Click(object sender, EventArgs e)
        {
            DataTable summaryDt = new DataTable();
            string searchTarget = _txtLogMismatchTarget.Text.Trim();
            int filterMode = Math.Max(0, _cmbLogMismatchFilter.SelectedIndex);

            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (_chkDemoMode.Checked)
                {
                    summaryDt.Columns.Add("차트번호");
                    summaryDt.Columns.Add("현재환자명");
                    summaryDt.Columns.Add("주민등록번호");
                    summaryDt.Columns.Add("로그환자명");
                    summaryDt.Columns.Add("로그원차트");
                    summaryDt.Columns.Add("백업환자명");
                    summaryDt.Columns.Add("백업원차트");
                    summaryDt.Columns.Add("백업확인건수", typeof(int));
                    summaryDt.Columns.Add("전체처방건수", typeof(int));
                    summaryDt.Columns.Add("로그확인건수", typeof(int));
                    summaryDt.Columns.Add("로그미확인건수", typeof(int));
                    summaryDt.Columns.Add("이상건수", typeof(int));
                    summaryDt.Columns.Add("청구환자명");
                    summaryDt.Columns.Add("청구차트");
                    summaryDt.Columns.Add("충돌유형");
                    summaryDt.Columns.Add("최근조제일");

                    // Demo Case 1: 송우연 (570201-2) vs 김영금 (570201-2)
                    summaryDt.Rows.Add(
                        "0000436221",
                        "김영금",
                        "570201-2******",
                        "김영금, 송우연",
                        "0000299486, 0000336491, 0000463116",
                        "김영금, 송우연, 김영심, 김인순, 한윤숙",
                        "0000290168, 0000299486, 0000290324, 0000291432, 0000313949",
                        161,
                        163,
                        12,
                        151,
                        12,
                        "김영금, 송우연",
                        "0000436221",
                        "다중 환자 병합 의심 (송우연 처방 혼입)",
                        "2026-03-13"
                    );

                    // Demo Case 2: 박복순 vs 천미선
                    summaryDt.Rows.Add(
                        "0000184791",
                        "박복순",
                        "590307-2******",
                        "천미선, 박복순",
                        "0000184791",
                        "천미선, 박복순",
                        "0000184791",
                        8,
                        20,
                        8,
                        12,
                        7,
                        "천미선",
                        "0000184791",
                        "처방로그 환자명 불일치",
                        "2026-06-05"
                    );

                    if (!string.IsNullOrEmpty(searchTarget))
                    {
                        for (int i = summaryDt.Rows.Count - 1; i >= 0; i--)
                        {
                            string chrt = Convert.ToString(summaryDt.Rows[i]["차트번호"]);
                            string nm = Convert.ToString(summaryDt.Rows[i]["현재환자명"]);
                            string logNm = Convert.ToString(summaryDt.Rows[i]["로그환자명"]);
                            string logChrt = Convert.ToString(summaryDt.Rows[i]["로그원차트"]);
                            string backupNm = Convert.ToString(summaryDt.Rows[i]["백업환자명"]);
                            string backupChrt = Convert.ToString(summaryDt.Rows[i]["백업원차트"]);
                            if (!chrt.Contains(searchTarget) && !nm.Contains(searchTarget) && !logNm.Contains(searchTarget) && !logChrt.Contains(searchTarget)
                                && !backupNm.Contains(searchTarget) && !backupChrt.Contains(searchTarget))
                            {
                                summaryDt.Rows.RemoveAt(i);
                            }
                        }
                    }
                }
                else if (filterMode == 7)
                {
                    LoadJuminEncryptionMismatchSummary(summaryDt, searchTarget);
                }
                else
                {
                    // Rank logs/claims once, then join one evidence row per prescription.
                    // This is substantially faster than repeating an OUTER APPLY lookup for
                    // every TBSID040_03 row on large production databases.
                    string backupDatabaseName = FindAttachedPrescriptionBackupDatabase();
                    string sql = @"
WITH LatestLog AS
(
    SELECT pl.PRESERIAL, pl.PANAME, pl.PANUM, pl.PRES_TEXT, pl.STATE_GUBUN, pl.INDATE,
           ROW_NUMBER() OVER (PARTITION BY pl.PRESERIAL ORDER BY pl.INDATE DESC) AS RN
    FROM PMPLUS_JOBLOG.dbo.PM_PRES_LOG pl WITH (NOLOCK)
    WHERE pl.STATE_GUBUN IN ('I', 'U')
),
LatestClaim AS
(
    SELECT ph.DRUG_SEQ, ph.CHRTNO, ph.PAT_NM, ph.PAT_JUMIN_NO, ph.INS_NUMBER,
           ROW_NUMBER() OVER (PARTITION BY ph.DRUG_SEQ ORDER BY ph.DRUG_SEQ) AS RN
    FROM dbo.TBSIB_H024_1 ph WITH (NOLOCK)
),
" + BuildPrescriptionBackupCte(backupDatabaseName) + @"
SELECT 
    r.CHRTNO,
    r.PAT_NM,
    r.PAT_JUMIN_NO,
    l.PANAME AS LOG_PANAME,
    l.PANUM AS LOG_PANUM,
    h.PAT_NM AS CLAIM_PAT_NM,
    h.CHRTNO AS CLAIM_CHRTNO,
    h.PAT_JUMIN_NO AS CLAIM_JUMIN_NO,
    h.INS_NUMBER AS CLAIM_INS_NUM,
    b.CHRTNO AS BACKUP_CHRTNO,
    b.PAT_NM AS BACKUP_PAT_NM,
    b.PAT_JUMIN_NO AS BACKUP_JUMIN_NO,
    r.DRUG_SEQ,
    r.PRES_DTIME,
    l.STATE_GUBUN,
    l.INDATE,
    parsed.LOG_CHRTNO,
    CASE WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL
                   AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(l.PANAME)) THEN 1 ELSE 0 END AS LOG_NAME_MISMATCH,
    CASE WHEN NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL
                   AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(h.PAT_NM)) THEN 1 ELSE 0 END AS CLAIM_NAME_MISMATCH,
    CASE WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL AND NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL
                   AND LTRIM(RTRIM(l.PANAME)) <> LTRIM(RTRIM(h.PAT_NM)) THEN 1 ELSE 0 END AS CLAIM_LOG_MISMATCH,
    CASE WHEN NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NOT NULL
                   AND LTRIM(RTRIM(ISNULL(r.CHRTNO, ''))) <> LTRIM(RTRIM(parsed.LOG_CHRTNO)) THEN 1 ELSE 0 END AS LOG_CHART_MISMATCH,
    CASE WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
                   AND (LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(ISNULL(b.PAT_NM, '')))
                     OR (NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(b.PAT_JUMIN_NO)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND LEFT(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), 7)
                             <> LEFT(REPLACE(REPLACE(LTRIM(RTRIM(b.PAT_JUMIN_NO)), '-', ''), ' ', ''), 7)))
              THEN 1 ELSE 0 END AS BACKUP_MISMATCH,
    CASE 
        WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
             AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(ISNULL(b.PAT_NM, ''))) THEN N'백업 원환자 불일치'
        WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(l.PANAME)) THEN N'처방로그 환자명 불일치'
        WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL AND NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL AND LTRIM(RTRIM(l.PANAME)) <> LTRIM(RTRIM(h.PAT_NM)) THEN N'청구/로그 환자 불일치'
        WHEN NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(h.PAT_NM)) THEN N'청구 환자명 불일치'
        WHEN NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NOT NULL AND LTRIM(RTRIM(ISNULL(r.CHRTNO, ''))) <> LTRIM(RTRIM(parsed.LOG_CHRTNO)) THEN N'로그 원차트 불일치'
        ELSE N'정상'
    END AS ISSUE_REASON
INTO #Evidence
FROM dbo.TBSID040_03 r WITH (NOLOCK)
LEFT JOIN LatestLog l ON l.PRESERIAL = r.DRUG_SEQ AND l.RN = 1
LEFT JOIN LatestClaim h ON h.DRUG_SEQ = r.DRUG_SEQ AND h.RN = 1
LEFT JOIN BackupPres b ON b.DRUG_SEQ = r.DRUG_SEQ
OUTER APPLY (SELECT CHARINDEX(N'#M#', l.PRES_TEXT) AS MARKER_POS) marker
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, marker.MARKER_POS + 3) AS P1) d1
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d1.P1 + 1) AS P2) d2
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d2.P2 + 1) AS P3) d3
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d3.P3 + 1) AS P4) d4
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d4.P4 + 1) AS P5) d5
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d5.P5 + 1) AS P6) d6
OUTER APPLY
(
    SELECT CASE WHEN marker.MARKER_POS > 0 AND d5.P5 > 0 AND d6.P6 > d5.P5
                THEN SUBSTRING(l.PRES_TEXT, d5.P5 + 1, d6.P6 - d5.P5 - 1)
                ELSE N'' END AS LOG_CHRTNO
) parsed;

CREATE INDEX IX_PMHELPER_EVIDENCE_CHRTNO ON #Evidence(CHRTNO);

SELECT 
    m.CHRTNO AS [차트번호],
    MAX(m.PAT_NM) AS [현재환자명],
    MAX(m.PAT_JUMIN_NO) AS [주민등록번호],
    STUFF((
        SELECT DISTINCT N', ' + l2.LOG_PANAME 
        FROM #Evidence l2
        WHERE l2.CHRTNO = m.CHRTNO AND l2.LOG_PANAME IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [로그환자명],
    STUFF((
        SELECT DISTINCT N', ' + e2.LOG_CHRTNO
        FROM #Evidence e2
        WHERE e2.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(e2.LOG_CHRTNO)), '') IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [로그원차트],
    STUFF((
        SELECT DISTINCT N', ' + b2.BACKUP_PAT_NM
        FROM #Evidence b2
        WHERE b2.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(b2.BACKUP_PAT_NM)), '') IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [백업환자명],
    STUFF((
        SELECT DISTINCT N', ' + b3.BACKUP_CHRTNO
        FROM #Evidence b3
        WHERE b3.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(b3.BACKUP_CHRTNO)), '') IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [백업원차트],
    SUM(CASE WHEN NULLIF(LTRIM(RTRIM(m.BACKUP_CHRTNO)), '') IS NOT NULL THEN 1 ELSE 0 END) AS [백업확인건수],
    COUNT(*) AS [전체처방건수],
    SUM(CASE WHEN NULLIF(LTRIM(RTRIM(m.LOG_PANAME)), '') IS NOT NULL THEN 1 ELSE 0 END) AS [로그확인건수],
    SUM(CASE WHEN NULLIF(LTRIM(RTRIM(m.LOG_PANAME)), '') IS NULL THEN 1 ELSE 0 END) AS [로그미확인건수],
    SUM(CASE WHEN m.BACKUP_MISMATCH = 1 OR m.LOG_NAME_MISMATCH = 1 OR m.CLAIM_NAME_MISMATCH = 1 OR m.CLAIM_LOG_MISMATCH = 1 THEN 1 ELSE 0 END) AS [이상건수],
    STUFF((
        SELECT DISTINCT N', ' + h2.CLAIM_PAT_NM 
        FROM #Evidence h2
        WHERE h2.CHRTNO = m.CHRTNO AND h2.CLAIM_PAT_NM IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [청구환자명],
    STUFF((
        SELECT DISTINCT N', ' + h3.CLAIM_CHRTNO
        FROM #Evidence h3
        WHERE h3.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(h3.CLAIM_CHRTNO)), '') IS NOT NULL
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS [청구차트],
    CASE
        WHEN MAX(m.BACKUP_MISMATCH) = 1 THEN N'백업 원환자 불일치'
        WHEN (SELECT COUNT(DISTINCT e3.LOG_PANAME) FROM #Evidence e3 WHERE e3.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(e3.LOG_PANAME)), '') IS NOT NULL) > 1 THEN N'다중 환자 병합 의심'
        WHEN MAX(m.LOG_NAME_MISMATCH) = 1 THEN N'처방로그 환자명 불일치'
        WHEN MAX(m.CLAIM_LOG_MISMATCH) = 1 THEN N'청구/로그 환자 불일치'
        WHEN MAX(m.CLAIM_NAME_MISMATCH) = 1 THEN N'청구 환자명 불일치'
        WHEN MAX(m.LOG_CHART_MISMATCH) = 1 THEN N'로그 원차트 불일치'
        ELSE N'확인 필요'
    END AS [충돌유형],
    MAX(m.PRES_DTIME) AS [최근조제일]
FROM #Evidence m
WHERE (@target = '' OR m.CHRTNO LIKE '%' + @target + '%' OR m.LOG_CHRTNO LIKE '%' + @target + '%'
    OR m.BACKUP_CHRTNO LIKE '%' + @target + '%' OR m.BACKUP_PAT_NM LIKE '%' + @target + '%'
    OR m.PAT_NM LIKE '%' + @target + '%' OR m.LOG_PANAME LIKE '%' + @target + '%' OR m.PAT_JUMIN_NO LIKE @target + '%')
GROUP BY m.CHRTNO
HAVING
       (@filterMode IN (0, 4) AND SUM(CASE WHEN m.BACKUP_MISMATCH = 1 OR m.LOG_NAME_MISMATCH = 1 OR m.CLAIM_NAME_MISMATCH = 1 OR m.CLAIM_LOG_MISMATCH = 1 THEN 1 ELSE 0 END) > 0)
    OR (@filterMode = 1 AND MAX(m.LOG_NAME_MISMATCH) = 1)
    OR (@filterMode = 2 AND (MAX(m.CLAIM_NAME_MISMATCH) = 1 OR MAX(m.CLAIM_LOG_MISMATCH) = 1))
    OR (@filterMode = 3 AND (SELECT COUNT(DISTINCT e4.LOG_PANAME) FROM #Evidence e4 WHERE e4.CHRTNO = m.CHRTNO AND NULLIF(LTRIM(RTRIM(e4.LOG_PANAME)), '') IS NOT NULL) > 1)
    OR (@filterMode = 5 AND SUM(CASE WHEN NULLIF(LTRIM(RTRIM(m.LOG_PANAME)), '') IS NULL THEN 1 ELSE 0 END) > 0)
    OR (@filterMode = 6 AND MAX(m.BACKUP_MISMATCH) = 1)
ORDER BY [최근조제일] DESC;

DROP TABLE #Evidence;";

                    using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 120;
                        cmd.Parameters.Add("@target", SqlDbType.NVarChar, 50).Value = searchTarget;
                        cmd.Parameters.Add("@filterMode", SqlDbType.Int).Value = filterMode;
                        conn.Open();
                        adapter.Fill(summaryDt);
                    }
                }

                bool juminEncryptionMode = filterMode == 7;
                if (juminEncryptionMode)
                {
                    _juminClassificationAll = summaryDt;
                    _juminClassificationView = "복구 대상";
                    UpdateJuminClassificationViewControls();
                    summaryDt = CreateJuminClassificationDisplayTable(_juminClassificationView);
                }
                else
                {
                    _juminClassificationAll = null;
                    if (_pnlJuminClassificationViews != null) _pnlJuminClassificationViews.Visible = false;
                    if (_dgvLogMismatchSummary.Columns.Contains("암호복구선택"))
                        _dgvLogMismatchSummary.Columns.Remove("암호복구선택");
                    _dgvLogMismatchSummary.ReadOnly = true;
                }

                _dgvLogMismatchSummary.DataSource = summaryDt;
                if (juminEncryptionMode) ConfigureJuminEncryptionSummaryGrid();
                else ApplyContentSizedColumns(_dgvLogMismatchSummary);
                _btnLogMismatchExport.Enabled = summaryDt.Rows.Count > 0;
                _btnLogRestoreSelectAll.Enabled = !juminEncryptionMode;
                _btnLogRestoreDeselectAll.Enabled = !juminEncryptionMode;
                _btnLogRestoreSelectAll.Text = juminEncryptionMode ? "✓ 현재 환자 전체선택" : "✓ 환자별 선택";
                _btnLogRestoreDeselectAll.Text = juminEncryptionMode ? "✕ 선택해제" : "✗ 해제";
                _cmbLogRestorePatientGroup.Enabled = !juminEncryptionMode;
                _txtLogRestoreNewChrtNo.Enabled = !juminEncryptionMode;
                _btnLogRestoreSeparate.Enabled = !juminEncryptionMode;
                _btnLogRestoreSeparate.Text = juminEncryptionMode
                    ? "🔐 현재 환자 암호문만 복구"
                    : "🛠️ 선택 처방을 원차트로 분리/복구";
                LayoutLogMismatchActionControls();

                if (juminEncryptionMode)
                {
                    int possible = _juminClassificationAll.Select("[복구가능] = '복구 가능'").Length;
                    int impossible = _juminClassificationAll.Select("[복구가능] = '복구 불가'").Length;
                    int noEvidence = _juminClassificationAll.Select("[복구가능] = '근거 없음'").Length;
                    int unidentified = _juminClassificationAll.Select("[복구가능] = '식별 불가'").Length;
                    int normal = _juminClassificationAll.Select("[복구가능] = '이미 정상'").Length;
                    _lblLogMismatchSummary.Text = string.Format(
                        "현재 표시: 복구 대상 {0}명 (가능 {1} / 불가 {2}) | 근거 없음 {3} / 식별 불가 {4} / 이미 정상 {5}",
                        possible + impossible, possible, impossible, noEvidence, unidentified, normal);
                    ShowToast(string.Format("복구 대상 {0}명 중 복구 가능 {1}명", possible + impossible, possible), ColorEmerald);
                }
                else
                {
                    _lblLogMismatchSummary.Text = string.Format("로그/청구 불일치 진단 완료: 총 {0}개 차트번호에서 이상 감지됨", summaryDt.Rows.Count);
                    ShowToast(string.Format("무결성 진단 완료: {0}개 차트 감지", summaryDt.Rows.Count), ColorEmerald);
                }

                if (summaryDt.Rows.Count > 0)
                {
                    LoadLogMismatchDetailForChart(Convert.ToString(summaryDt.Rows[0]["차트번호"]));
                    if (juminEncryptionMode)
                    {
                        UpdateJuminRestoreButtonState();
                    }
                }
                else
                {
                    _dgvLogMismatchDetail.DataSource = null;
                    _lblLogMismatchDetailInfo.Text = juminEncryptionMode
                        ? "표시할 활성 고객이 없습니다."
                        : "감지된 이상 내역이 없습니다. (모든 조제 기록이 로그 및 청구 데이터와 정상 일치)";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("로그/청구 기반 무결성 검사 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private DataTable CreateJuminClassificationDisplayTable(string viewName)
        {
            if (_juminClassificationAll == null) return new DataTable();

            string expression;
            if (viewName == "근거 없음") expression = "[복구가능] = '근거 없음'";
            else if (viewName == "식별 불가") expression = "[복구가능] = '식별 불가'";
            else expression = "[복구가능] = '복구 가능' OR [복구가능] = '복구 불가'";

            DataTable display = _juminClassificationAll.Clone();
            foreach (DataRow row in _juminClassificationAll.Select(expression))
            {
                display.ImportRow(row);
            }
            return display;
        }

        private void UpdateJuminClassificationViewControls()
        {
            if (_pnlJuminClassificationViews == null) return;

            bool available = _juminClassificationAll != null && _juminClassificationAll.Columns.Contains("복구가능");
            _pnlJuminClassificationViews.Visible = available;
            if (!available) return;

            int possible = _juminClassificationAll.Select("[복구가능] = '복구 가능'").Length;
            int impossible = _juminClassificationAll.Select("[복구가능] = '복구 불가'").Length;
            int noEvidence = _juminClassificationAll.Select("[복구가능] = '근거 없음'").Length;
            int unidentified = _juminClassificationAll.Select("[복구가능] = '식별 불가'").Length;
            int normal = _juminClassificationAll.Select("[복구가능] = '이미 정상'").Length;

            _btnJuminShowRestoreTargets.Text = string.Format("복구대상 {0:N0}명", possible + impossible);
            _btnJuminShowNoEvidence.Text = string.Format("근거없음 {0:N0}명", noEvidence);
            _btnJuminShowUnidentified.Text = string.Format("식별불가 {0:N0}명", unidentified);
            _lblJuminNormalCount.Text = string.Format("이미정상 {0:N0}명", normal);
            ResizeJuminClassificationViewControls();

            _btnJuminShowRestoreTargets.Enabled = possible + impossible > 0;
            _btnJuminShowNoEvidence.Enabled = noEvidence > 0;
            _btnJuminShowUnidentified.Enabled = unidentified > 0;
            _btnJuminShowRestoreTargets.BackColor = _juminClassificationView == "복구 대상" ? ColorIndigo : ColorBgCard;
            _btnJuminShowNoEvidence.BackColor = _juminClassificationView == "근거 없음" ? ColorIndigo : ColorBgCard;
            _btnJuminShowUnidentified.BackColor = _juminClassificationView == "식별 불가" ? ColorIndigo : ColorBgCard;
        }

        private void ResizeJuminClassificationViewControls()
        {
            if (_pnlJuminClassificationViews == null || _btnJuminShowRestoreTargets == null ||
                _btnJuminShowNoEvidence == null || _btnJuminShowUnidentified == null ||
                _lblJuminNormalCount == null) return;

            Control[] controls = { _btnJuminShowRestoreTargets, _btnJuminShowNoEvidence, _btnJuminShowUnidentified, _lblJuminNormalCount };
            int totalWidth = 0;
            foreach (Control control in controls)
            {
                int textWidth = TextRenderer.MeasureText(control.Text ?? "", control.Font,
                    new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width;
                int minimumWidth = control == _lblJuminNormalCount ? 145 : 135;
                control.Width = Math.Max(minimumWidth, textWidth + 44);
                totalWidth += control.Width + control.Margin.Horizontal;
            }

            // A very narrow left pane may still require horizontal scrolling, but never clip
            // the count inside the button itself.
            _pnlJuminClassificationViews.AutoScroll = totalWidth > _pnlJuminClassificationViews.ClientSize.Width;
        }

        private void ShowJuminClassificationView(string viewName)
        {
            if (_juminClassificationAll == null || _cmbLogMismatchFilter == null ||
                _cmbLogMismatchFilter.SelectedIndex != 7) return;

            _juminClassificationView = viewName;
            UpdateJuminClassificationViewControls();
            DataTable display = CreateJuminClassificationDisplayTable(viewName);
            _dgvLogMismatchSummary.DataSource = display;
            ConfigureJuminEncryptionSummaryGrid();
            _btnLogMismatchExport.Enabled = display.Rows.Count > 0;

            _lblLogMismatchSummary.Text = string.Format(
                "현재 표시: {0} {1:N0}명 | 복구 대상·근거 없음·식별 불가 버튼으로 목록을 전환합니다. 이미 정상은 건수만 표시합니다.",
                viewName, display.Rows.Count);
            if (display.Rows.Count > 0)
            {
                string chartNo = Convert.ToString(display.Rows[0]["차트번호"]);
                LoadLogMismatchDetailForChart(chartNo);
            }
            else
            {
                _dgvLogMismatchDetail.DataSource = null;
            }
            UpdateJuminRestoreButtonState();
        }

        private void ConfigureJuminEncryptionSummaryGrid()
        {
            if (_dgvLogMismatchSummary == null || _dgvLogMismatchSummary.Columns.Count == 0) return;

            if (_dgvLogMismatchSummary.Columns.Contains("암호복구선택"))
                _dgvLogMismatchSummary.Columns.Remove("암호복구선택");
            _dgvLogMismatchSummary.ReadOnly = true;
            foreach (DataGridViewColumn column in _dgvLogMismatchSummary.Columns)
            {
                column.ReadOnly = true;
            }

            string[] hiddenColumns =
            {
                "로그환자명", "로그원차트", "로그확인건수", "로그미확인건수",
                "이상건수", "청구환자명", "청구차트", "백업환자명"
            };
            foreach (string columnName in hiddenColumns)
            {
                if (_dgvLogMismatchSummary.Columns.Contains(columnName))
                    _dgvLogMismatchSummary.Columns[columnName].Visible = false;
            }

            Dictionary<string, int> widths = new Dictionary<string, int>
            {
                { "차트번호", 105 }, { "현재환자명", 120 }, { "주민등록번호", 125 },
                { "백업원차트", 190 }, { "백업확인건수", 90 }, { "전체처방건수", 90 },
                { "복구가능", 95 }, { "선행복구필요건수", 115 }, { "충돌유형", 300 }, { "최근조제일", 120 },
                { "현재암호길이", 95 }, { "백업암호길이", 95 }
            };
            _dgvLogMismatchSummary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            foreach (KeyValuePair<string, int> item in widths)
            {
                if (!_dgvLogMismatchSummary.Columns.Contains(item.Key)) continue;
                DataGridViewColumn column = _dgvLogMismatchSummary.Columns[item.Key];
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                column.MinimumWidth = Math.Min(60, item.Value);
                column.Width = item.Value;
            }

            foreach (DataGridViewRow row in _dgvLogMismatchSummary.Rows)
            {
                if (row.IsNewRow) continue;
                string status = Convert.ToString(row.Cells["복구가능"].Value);
                if (status == "복구 가능")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(20, 64, 45);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(167, 243, 208);
                }
                else if (status == "복구 불가")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(64, 20, 20);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(254, 202, 202);
                }
                else if (status == "근거 없음" || status == "식별 불가")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(55, 48, 25);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(253, 230, 138);
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(23, 37, 56);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(147, 197, 253);
                }
            }
            UpdateJuminRestoreButtonState();
        }

        private List<DataGridViewRow> GetCheckedJuminRestoreRows()
        {
            List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
            if (_dgvLogMismatchSummary == null || _dgvLogMismatchSummary.CurrentRow == null ||
                _dgvLogMismatchDetail == null)
                return selectedRows;

            DataGridViewRow currentPatientRow = _dgvLogMismatchSummary.CurrentRow;
            string status = Convert.ToString(currentPatientRow.Cells["복구가능"].Value);
            int totalDetailRows = _dgvLogMismatchDetail.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
            int selectedDetailRows = CountSelectedJuminDetailRows();
            if (status == "복구 가능" && totalDetailRows > 0 && selectedDetailRows == totalDetailRows)
                selectedRows.Add(currentPatientRow);
            return selectedRows;
        }

        private int CountSelectedJuminDetailRows()
        {
            if (_dgvLogMismatchDetail == null || !_dgvLogMismatchDetail.Columns.Contains("선택")) return 0;
            int count = 0;
            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                if (row.IsNewRow) continue;
                if (row.Cells["선택"].Value != null && Convert.ToBoolean(row.Cells["선택"].Value)) count++;
            }
            return count;
        }

        private void SetJuminDetailChecked(bool check)
        {
            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Rows.Count == 0 ||
                _dgvLogMismatchSummary == null || _dgvLogMismatchSummary.CurrentRow == null) return;

            bool patientRestorable = Convert.ToString(
                _dgvLogMismatchSummary.CurrentRow.Cells["복구가능"].Value) == "복구 가능";
            if (check && !patientRestorable)
            {
                MessageBox.Show("현재 환자는 암호문 자동 복구 가능 상태가 아닙니다.", "복구 불가",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int count = 0;
            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                if (row.IsNewRow || row.Cells["선택"] == null) continue;
                row.Cells["선택"].Value = check && patientRestorable;
                if (check && patientRestorable) count++;
            }
            UpdateJuminRestoreButtonState();
            ShowToast(check ? string.Format("현재 환자의 처방 {0:N0}건을 복구 확인 대상으로 선택했습니다.", count) : "현재 환자의 선택을 모두 해제했습니다.",
                check ? ColorEmerald : ColorTextSec);
        }

        private void UpdateJuminRestoreButtonState()
        {
            if (_btnLogRestoreSeparate == null || _cmbLogMismatchFilter == null ||
                _cmbLogMismatchFilter.SelectedIndex != 7) return;

            bool patientRestorable = _dgvLogMismatchSummary != null &&
                _dgvLogMismatchSummary.CurrentRow != null &&
                Convert.ToString(_dgvLogMismatchSummary.CurrentRow.Cells["복구가능"].Value) == "복구 가능";
            int totalDetailRows = _dgvLogMismatchDetail == null
                ? 0
                : _dgvLogMismatchDetail.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow);
            int selectedCount = CountSelectedJuminDetailRows();
            bool allSelected = totalDetailRows > 0 && selectedCount == totalDetailRows;
            _btnLogRestoreSelectAll.Enabled = patientRestorable && totalDetailRows > 0;
            _btnLogRestoreDeselectAll.Enabled = selectedCount > 0;
            _btnLogRestoreSeparate.Enabled = patientRestorable && allSelected;
            _btnLogRestoreSeparate.Text = string.Format("🔐 현재 환자 암호문만 복구 ({0:N0}/{1:N0}건)", selectedCount, totalDetailRows);
            LayoutLogMismatchActionControls();
        }

        private void DgvLogMismatchSummary_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || _dgvLogMismatchSummary.DataSource == null) return;
            string chrtNo = Convert.ToString(_dgvLogMismatchSummary.Rows[e.RowIndex].Cells["차트번호"].Value);
            LoadLogMismatchDetailForChart(chrtNo);
        }

        private void DgvLogMismatchSummary_SelectionChanged(object sender, EventArgs e)
        {
            if (_dgvLogMismatchSummary.CurrentRow == null || _dgvLogMismatchSummary.DataSource == null) return;
            string chrtNo = Convert.ToString(_dgvLogMismatchSummary.CurrentRow.Cells["차트번호"].Value);
            LoadLogMismatchDetailForChart(chrtNo);
            UpdateJuminRestoreButtonState();
        }

        private void LoadLogMismatchDetailForChart(string chrtNo)
        {
            if (string.IsNullOrEmpty(chrtNo)) return;

            DataTable detailDt = new DataTable();
            string suggestedRestoreChart = "";
            string currentPatName = "";

            try
            {
                if (_chkDemoMode.Checked)
                {
                    detailDt.Columns.Add("선택", typeof(bool));
                    detailDt.Columns.Add("복구가능", typeof(bool));
                    detailDt.Columns.Add("조제번호");
                    detailDt.Columns.Add("조제일자");
                    detailDt.Columns.Add("원장차트");
                    detailDt.Columns.Add("원장환자명");
                    detailDt.Columns.Add("복구근거");
                    detailDt.Columns.Add("복구원차트");
                    detailDt.Columns.Add("복구환자명");
                    detailDt.Columns.Add("복구주민번호");
                    detailDt.Columns.Add("백업원차트");
                    detailDt.Columns.Add("백업환자명");
                    detailDt.Columns.Add("백업주민번호");
                    detailDt.Columns.Add("로그환자명");
                    detailDt.Columns.Add("로그주민번호");
                    detailDt.Columns.Add("로그원차트");
                    detailDt.Columns.Add("청구차트");
                    detailDt.Columns.Add("청구환자명");
                    detailDt.Columns.Add("청구주민번호");
                    detailDt.Columns.Add("청구증번호");
                    detailDt.Columns.Add("진단상태");
                    detailDt.Columns.Add("로그원문요약");

                    if (chrtNo == "0000436221")
                    {
                        currentPatName = "김영금";
                        suggestedRestoreChart = ""; // 여러 원차트가 있으므로 한 묶음씩 선택

                        // Song Woo-yeon rows merged into Kim Young-geum
                        detailDt.Rows.Add(true, true, "20260313000153", "2026-03-13", "0000436221", "김영금", "백업", "0000299486", "송우연", "570201-2******", "0000299486", "송우연", "570201-2******", "송우연", "570201-2******", "0000299486", "0000436221", "김영금", "570201-2******", "27027675887", "🔴 백업 원본 확인", "#M#57|20260313|...|0|0|0000299486|...");
                        detailDt.Rows.Add(true, true, "20260212000156", "2026-02-12", "0000436221", "김영금", "백업", "0000299486", "송우연", "570201-2******", "0000299486", "송우연", "570201-2******", "송우연", "570201-2******", "0000299486", "0000436221", "김영금", "570201-2******", "27027675887", "🔴 백업 원본 확인", "#M#62|20260212|...|4|E|0000299486|...");
                        detailDt.Rows.Add(true, true, "20260716000320", "2026-07-16", "0000436221", "김영금", "처방로그", "0000336491", "송우연", "570201-2******", "", "", "", "송우연", "570201-2******", "0000336491", "0000436221", "김영금", "570201-2******", "27027675887", "⚠️ 로그/청구 환자 상호 불일치", "#M#121|20260716|...|0|0|0000336491|...");
                        detailDt.Rows.Add(false, false, "20260729000090", "2026-07-29", "0000436221", "김영금", "처방로그", "0000326751", "김영금", "570201-2******", "", "", "", "김영금", "570201-2******", "0000326751", "0000436221", "김영금", "570201-2******", "80835027829", "⚠️ 로그 원차트 불일치", "#M#50|20260729|...|0|0|0000326751|...");
                    }
                    else
                    {
                        currentPatName = "박복순";
                        suggestedRestoreChart = "0000184791";
                        detailDt.Rows.Add(true, true, "2026060500001", "2026-06-05", chrtNo, "박복순", "처방로그", "0000184791", "천미선", "770315-2******", "", "", "", "천미선", "770315-2******", "0000184791", chrtNo, "박복순", "590307-2******", "12345678901", "⚠️ 로그/청구 환자 상호 불일치", "#M#10|20260605|...|0|0|0000184791|...");
                        detailDt.Rows.Add(false, false, "2026060100002", "2026-06-01", chrtNo, "박복순", "처방로그", chrtNo, "박복순", "590307-2******", "", "", "", "박복순", "590307-2******", chrtNo, chrtNo, "박복순", "590307-2******", "98765432101", "정상 일치", "#M#9|20260601|...|0|0|0000184791|...");
                    }
                }
                else
                {
                    string backupDatabaseName = FindAttachedPrescriptionBackupDatabase();
                    string sql = @"
WITH LatestLog AS
(
    SELECT pl.PRESERIAL, pl.PANAME, pl.PANUM, pl.PRES_TEXT, pl.STATE_GUBUN, pl.INDATE,
           ROW_NUMBER() OVER (PARTITION BY pl.PRESERIAL ORDER BY pl.INDATE DESC) AS RN
    FROM PMPLUS_JOBLOG.dbo.PM_PRES_LOG pl WITH (NOLOCK)
    WHERE pl.STATE_GUBUN IN ('I', 'U')
),
LatestClaim AS
(
    SELECT ph.DRUG_SEQ, ph.CHRTNO, ph.PAT_NM, ph.PAT_JUMIN_NO, ph.INS_NUMBER,
           ROW_NUMBER() OVER (PARTITION BY ph.DRUG_SEQ ORDER BY ph.DRUG_SEQ) AS RN
    FROM dbo.TBSIB_H024_1 ph WITH (NOLOCK)
),
" + BuildPrescriptionBackupCte(backupDatabaseName) + @"
SELECT 
    CAST(CASE WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(b.CHRTNO)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(b.PAT_NM)), '') IS NOT NULL
                   AND identityCheck.BACKUP_IDENTITY_MISMATCH = 1
              THEN 1
              WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NULL
                   AND NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL
                   AND identityCheck.LOG_IDENTITY_MISMATCH = 1
              THEN 1 ELSE 0 END AS BIT) AS [선택],
    CAST(CASE WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(b.CHRTNO)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(b.PAT_NM)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(b.PAT_JUMIN_NO)), '') IS NOT NULL
                   AND identityCheck.BACKUP_IDENTITY_MISMATCH = 1
              THEN 1
              WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NULL
                   AND NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL
                   AND NULLIF(LTRIM(RTRIM(l.PANUM)), '') IS NOT NULL
                   AND identityCheck.LOG_IDENTITY_MISMATCH = 1
              THEN 1 ELSE 0 END AS BIT) AS [복구가능],
    r.DRUG_SEQ AS [조제번호],
    r.PRES_DTIME AS [조제일자],
    r.CHRTNO AS [원장차트],
    r.PAT_NM AS [원장환자명],
    CASE WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL THEN N'백업(' + @backupDbDisplay + N')'
         WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL THEN N'처방로그'
         ELSE N'근거없음' END AS [복구근거],
    COALESCE(NULLIF(LTRIM(RTRIM(b.CHRTNO)), ''), NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), ''), N'') AS [복구원차트],
    COALESCE(NULLIF(LTRIM(RTRIM(b.PAT_NM)), ''), NULLIF(LTRIM(RTRIM(l.PANAME)), ''), N'') AS [복구환자명],
    COALESCE(NULLIF(LTRIM(RTRIM(b.PAT_JUMIN_NO)), ''), NULLIF(LTRIM(RTRIM(l.PANUM)), ''), N'') AS [복구주민번호],
    ISNULL(b.CHRTNO, N'') AS [백업원차트],
    ISNULL(b.PAT_NM, N'') AS [백업환자명],
    ISNULL(b.PAT_JUMIN_NO, N'') AS [백업주민번호],
    CASE WHEN l.PRESERIAL IS NULL THEN N'(로그없음)' ELSE ISNULL(NULLIF(LTRIM(RTRIM(l.PANAME)), ''), N'(환자명없음)') END AS [로그환자명],
    CASE WHEN l.PRESERIAL IS NULL THEN N'(로그없음)' ELSE ISNULL(NULLIF(LTRIM(RTRIM(l.PANUM)), ''), N'(주민번호없음)') END AS [로그주민번호],
    CASE WHEN l.PRESERIAL IS NULL THEN N'(로그없음)' ELSE ISNULL(NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), ''), N'(원차트없음)') END AS [로그원차트],
    ISNULL(h.CHRTNO, N'') AS [청구차트],
    ISNULL(h.PAT_NM, N'(청구없음)') AS [청구환자명],
    ISNULL(h.PAT_JUMIN_NO, N'') AS [청구주민번호],
    ISNULL(h.INS_NUMBER, N'') AS [청구증번호],
    CASE 
        WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
             AND identityCheck.BACKUP_IDENTITY_MISMATCH = 1
             THEN N'🔴 백업 원환자 확인 (' + ISNULL(b.CHRTNO,N'') + N' / ' + ISNULL(b.PAT_NM,N'') + N')'
        WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
             AND LTRIM(RTRIM(ISNULL(r.CHRTNO, ''))) <> LTRIM(RTRIM(ISNULL(b.CHRTNO, '')))
             THEN N'ℹ️ 동일 환자의 과거 차트번호 (' + ISNULL(b.CHRTNO,N'') + N') - 복구대상 아님'
        WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL THEN N'백업과 정상 일치'
        WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NULL THEN N'⚪ 로그없음 - 자동판정불가'
        WHEN NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NULL THEN N'⚪ 로그 원차트 없음 - 자동판정불가'
        WHEN identityCheck.LOG_IDENTITY_MISMATCH = 1 THEN N'⚠️ 로그 원환자 불일치 (' + l.PANAME + N')'
        WHEN NULLIF(LTRIM(RTRIM(parsed.LOG_CHRTNO)), '') IS NOT NULL
             AND LTRIM(RTRIM(ISNULL(r.CHRTNO, ''))) <> LTRIM(RTRIM(parsed.LOG_CHRTNO))
             THEN N'ℹ️ 동일 환자의 로그 과거 차트번호 (' + parsed.LOG_CHRTNO + N') - 복구대상 아님'
        WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL AND NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL
             AND LTRIM(RTRIM(l.PANAME)) <> LTRIM(RTRIM(h.PAT_NM)) THEN N'⚠️ 로그/청구 환자 상호 불일치'
        WHEN NULLIF(LTRIM(RTRIM(h.PAT_NM)), '') IS NOT NULL AND LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(h.PAT_NM)) THEN N'⚠️ 청구 환자명 불일치 (' + h.PAT_NM + N')'
        ELSE N'정상 일치'
    END AS [진단상태],
    ISNULL(SUBSTRING(l.PRES_TEXT, 1, 150), N'') AS [로그원문요약]
FROM dbo.TBSID040_03 r WITH (NOLOCK)
LEFT JOIN LatestLog l ON l.PRESERIAL = r.DRUG_SEQ AND l.RN = 1
LEFT JOIN LatestClaim h ON h.DRUG_SEQ = r.DRUG_SEQ AND h.RN = 1
LEFT JOIN BackupPres b ON b.DRUG_SEQ = r.DRUG_SEQ
OUTER APPLY (SELECT CHARINDEX(N'#M#', l.PRES_TEXT) AS MARKER_POS) marker
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, marker.MARKER_POS + 3) AS P1) d1
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d1.P1 + 1) AS P2) d2
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d2.P2 + 1) AS P3) d3
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d3.P3 + 1) AS P4) d4
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d4.P4 + 1) AS P5) d5
OUTER APPLY (SELECT CHARINDEX(N'|', l.PRES_TEXT, d5.P5 + 1) AS P6) d6
OUTER APPLY
(
    SELECT CASE WHEN marker.MARKER_POS > 0 AND d5.P5 > 0 AND d6.P6 > d5.P5
                THEN SUBSTRING(l.PRES_TEXT, d5.P5 + 1, d6.P6 - d5.P5 - 1)
                ELSE N'' END AS LOG_CHRTNO
) parsed
OUTER APPLY
(
    SELECT
        CASE WHEN NULLIF(LTRIM(RTRIM(b.DRUG_SEQ)), '') IS NOT NULL
                   AND (LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(ISNULL(b.PAT_NM, '')))
                     OR (NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(b.PAT_JUMIN_NO)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND LEFT(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), 7)
                             <> LEFT(REPLACE(REPLACE(LTRIM(RTRIM(b.PAT_JUMIN_NO)), '-', ''), ' ', ''), 7)))
             THEN 1 ELSE 0 END AS BACKUP_IDENTITY_MISMATCH,
        CASE WHEN NULLIF(LTRIM(RTRIM(l.PANAME)), '') IS NOT NULL
                   AND (LTRIM(RTRIM(ISNULL(r.PAT_NM, ''))) <> LTRIM(RTRIM(ISNULL(l.PANAME, '')))
                     OR (NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND NULLIF(REPLACE(REPLACE(LTRIM(RTRIM(l.PANUM)), '-', ''), ' ', ''), '') IS NOT NULL
                         AND LEFT(REPLACE(REPLACE(LTRIM(RTRIM(r.PAT_JUMIN_NO)), '-', ''), ' ', ''), 7)
                             <> LEFT(REPLACE(REPLACE(LTRIM(RTRIM(l.PANUM)), '-', ''), ' ', ''), 7)))
             THEN 1 ELSE 0 END AS LOG_IDENTITY_MISMATCH
) identityCheck
WHERE r.CHRTNO = @chrtNo
ORDER BY r.PRES_DTIME DESC;";

                    using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 60;
                        cmd.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = chrtNo;
                        cmd.Parameters.Add("@backupDbDisplay", SqlDbType.NVarChar, 128).Value = string.IsNullOrEmpty(backupDatabaseName) ? "없음" : backupDatabaseName;
                        conn.Open();
                        adapter.Fill(detailDt);
                    }

                    if (detailDt.Rows.Count > 0)
                    {
                        currentPatName = Convert.ToString(detailDt.Rows[0]["원장환자명"]);

                        // Suggest the evidence-backed original chart only when all initially
                        // selected rows agree. Multiple charts are restored one group at a time.
                        HashSet<string> originalCharts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (DataRow r in detailDt.Rows)
                        {
                            if (Convert.ToBoolean(r["선택"]))
                            {
                                string originalChart = Convert.ToString(r["복구원차트"]).Trim();
                                if (!string.IsNullOrEmpty(originalChart)) originalCharts.Add(originalChart);
                            }
                        }
                        if (originalCharts.Count == 1) suggestedRestoreChart = originalCharts.First();
                    }
                }

                bool juminEncryptionMode = _cmbLogMismatchFilter != null &&
                    _cmbLogMismatchFilter.SelectedIndex == 7;
                bool juminPatientRestorable = juminEncryptionMode &&
                    _dgvLogMismatchSummary != null && _dgvLogMismatchSummary.CurrentRow != null &&
                    Convert.ToString(_dgvLogMismatchSummary.CurrentRow.Cells["복구가능"].Value) == "복구 가능";
                if (juminEncryptionMode)
                {
                    string currentSummaryDigits = _dgvLogMismatchSummary != null &&
                        _dgvLogMismatchSummary.CurrentRow != null
                        ? Regex.Replace(Convert.ToString(_dgvLogMismatchSummary.CurrentRow.Cells["주민등록번호"].Value) ?? "", @"[^0-9]", "")
                        : "";
                    foreach (DataRow row in detailDt.Rows)
                    {
                        row["선택"] = false;
                        row["복구가능"] = juminPatientRestorable;

                        string evidence = Convert.ToString(row["복구근거"]).Trim();
                        string evidencePatient = Convert.ToString(row["복구환자명"]).Trim();
                        string currentPatient = Convert.ToString(row["원장환자명"]).Trim();
                        string evidenceDigits = Regex.Replace(Convert.ToString(row["복구주민번호"]) ?? "", @"[^0-9]", "");
                        bool sameIdentity = !string.IsNullOrEmpty(evidencePatient) &&
                            string.Equals(evidencePatient, currentPatient, StringComparison.OrdinalIgnoreCase) &&
                            evidenceDigits.Length >= 7 && currentSummaryDigits.Length >= 7 &&
                            string.Equals(evidenceDigits.Substring(0, 7), currentSummaryDigits.Substring(0, 7), StringComparison.Ordinal);

                        if (string.Equals(evidence, "근거없음", StringComparison.OrdinalIgnoreCase))
                        {
                            row["복구근거"] = "차트근거 없음(암호문 복구와 무관)";
                            row["복구원차트"] = "(확인불가·변경안함)";
                        }
                        else if (sameIdentity && evidence.StartsWith("백업", StringComparison.OrdinalIgnoreCase))
                        {
                            row["복구근거"] = "백업(동일환자 과거차트 참고)";
                        }
                        else if (sameIdentity && evidence.StartsWith("처방로그", StringComparison.OrdinalIgnoreCase))
                        {
                            row["복구근거"] = "처방로그(동일환자 과거차트 참고)";
                        }
                    }
                }

                _dgvLogMismatchDetail.DataSource = detailDt;
                if (!juminEncryptionMode) PopulateRestorePatientGroups(detailDt);
                ConfigureLogMismatchDetailHeaders(juminEncryptionMode);
                foreach (DataGridViewColumn column in _dgvLogMismatchDetail.Columns)
                {
                    column.ReadOnly = !string.Equals(column.Name, "선택", StringComparison.OrdinalIgnoreCase);
                }
                ApplyContentSizedColumns(_dgvLogMismatchDetail);
                FormatLogMismatchDetailGrid();

                int mismatchCount = 0;
                int logEvidenceCount = 0;
                int backupEvidenceCount = 0;
                int noEvidenceCount = 0;
                foreach (DataRow r in detailDt.Rows)
                {
                    if (Convert.ToBoolean(r["선택"])) mismatchCount++;
                    string logName = Convert.ToString(r["로그환자명"]);
                    if (!string.IsNullOrEmpty(logName) && logName != "(로그없음)") logEvidenceCount++;
                    string backupChart = Convert.ToString(r["백업원차트"]).Trim();
                    if (!string.IsNullOrEmpty(backupChart)) backupEvidenceCount++;
                    string restoreChart = Convert.ToString(r["복구원차트"]).Trim();
                    if (string.IsNullOrEmpty(restoreChart)) noEvidenceCount++;
                }
                _lblLogMismatchDetailInfo.Text = string.Format(
                    "차트 [{0}] {1}님: 전체 {2} / 백업확인 {3} / 로그확인 {4} / 근거없음 {5} / 확정이상 {6}",
                    chrtNo, currentPatName, detailDt.Rows.Count, backupEvidenceCount, logEvidenceCount, noEvidenceCount, mismatchCount);
                if (noEvidenceCount > 0)
                {
                    _lblLogMismatchDetailInfo.Text += " (백업·로그가 모두 없는 처방은 자동판정 불가)";
                }

                HashSet<string> detectedOriginalCharts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (DataRow r in detailDt.Rows)
                {
                    if (!Convert.ToBoolean(r["선택"])) continue;
                    string originalChart = Convert.ToString(r["복구원차트"]).Trim();
                    if (!string.IsNullOrEmpty(originalChart)) detectedOriginalCharts.Add(originalChart);
                }
                if (detectedOriginalCharts.Count > 1)
                {
                    foreach (DataRow r in detailDt.Rows) r["선택"] = false;
                    _lblLogMismatchDetailInfo.Text += string.Format(
                        " / 복구 원차트 {0}개 감지: 복구환자를 고른 뒤 '환자별 선택' 사용",
                        detectedOriginalCharts.Count);
                }

                if (!string.IsNullOrEmpty(suggestedRestoreChart))
                {
                    _txtLogRestoreNewChrtNo.Text = suggestedRestoreChart;
                }
                else
                {
                    _txtLogRestoreNewChrtNo.Text = "";
                }

                if (juminEncryptionMode)
                {
                    _lblLogMismatchDetailInfo.Text = juminPatientRestorable
                        ? string.Format("차트 [{0}] {1}님: 차트번호·환자명·처방내용은 변경하지 않음. 오른쪽 {2:N0}건 전체선택 시 관련 테이블의 암호문만 복구합니다.",
                            chrtNo, currentPatName, detailDt.Rows.Count)
                        : string.Format("차트 [{0}] {1}님은 자동 복구 불가 상태입니다. 오른쪽 처방은 참고용입니다.",
                            chrtNo, currentPatName);
                    _txtLogRestoreNewChrtNo.Text = "";
                    UpdateJuminRestoreButtonState();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("상세 내역 로드 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormatLogMismatchDetailGrid()
        {
            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Rows.Count == 0) return;

            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                string status = Convert.ToString(row.Cells["진단상태"].Value);
                if (status.StartsWith("🔴") || status.StartsWith("⚠️"))
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(64, 20, 20); // Dark Red
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(254, 202, 202); // Light Red text
                }
                else if (status.StartsWith("⚪"))
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(55, 48, 25);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(253, 230, 138);
                }
                else if (status.StartsWith("ℹ️"))
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(23, 37, 56);
                    row.DefaultCellStyle.ForeColor = Color.FromArgb(147, 197, 253);
                }
                else
                {
                    row.DefaultCellStyle.BackColor = ColorBgCard;
                    row.DefaultCellStyle.ForeColor = ColorTextMain;
                }
            }
        }

        private List<JuminRestoreTablePlan> CreateJuminRestoreTablePlans()
        {
            return new List<JuminRestoreTablePlan>
            {
                new JuminRestoreTablePlan
                {
                    TableName = "TBSIT000_01", DisplayName = "환자 마스터", ChartColumn = "CHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(50), CHRTNO), N'') + N'|' + ISNULL(CONVERT(nvarchar(50), PAT_SEQ), N'')"
                },
                new JuminRestoreTablePlan
                {
                    TableName = "TBSID040_03", DisplayName = "처방·조제", ChartColumn = "CHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "PAT_JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(100), DRUG_SEQ), N'')"
                },
                new JuminRestoreTablePlan
                {
                    TableName = "TBSIB_H024_1", DisplayName = "청구", ChartColumn = "CHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "PAT_JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(100), BILL_NO), N'') + N'|' + ISNULL(CONVERT(nvarchar(100), SPEC_SEQ_NO), N'')"
                },
                new JuminRestoreTablePlan
                {
                    TableName = "TBSIT000_02", DisplayName = "고객 메모", ChartColumn = "CHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(100), MEMO_IDX), N'')"
                },
                new JuminRestoreTablePlan
                {
                    TableName = "TEMP_MAPPING_CHRTNO", DisplayName = "차트 매핑", ChartColumn = "CHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(50), CHRTNO), N'') + N'|' + ISNULL(CONVERT(nvarchar(100), PAT_NM), N'') + N'|' + ISNULL(CONVERT(nvarchar(30), EXEDATE, 121), N'')"
                },
                new JuminRestoreTablePlan
                {
                    TableName = "TEMP_MAPPING_CHRTNO_SUB", DisplayName = "차트 매핑 상세", ChartColumn = "NEWCHRTNO",
                    NameColumn = "PAT_NM", JuminColumn = "JUMIN_NO",
                    RowKeyExpression = "ISNULL(CONVERT(nvarchar(50), CHRTNO), N'') + N'|' + ISNULL(CONVERT(nvarchar(50), NEWCHRTNO), N'') + N'|' + ISNULL(CONVERT(nvarchar(30), EXEDATE, 121), N'')"
                }
            };
        }

        private string ReadSingleJuminCipher(SqlConnection conn, SqlTransaction trans, string sql,
            string chartNo, string patientName, string juminPrefix, string sourceDescription, bool allowEmpty)
        {
            List<string> values = new List<string>();
            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            {
                cmd.CommandTimeout = 120;
                cmd.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = chartNo;
                cmd.Parameters.Add("@patNm", SqlDbType.NVarChar, 100).Value = patientName;
                cmd.Parameters.Add("@prefix", SqlDbType.NVarChar, 7).Value = juminPrefix;
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string value = reader.IsDBNull(0) ? "" : Convert.ToString(reader.GetValue(0)).Trim();
                        if ((allowEmpty || !string.IsNullOrEmpty(value)) && !values.Contains(value)) values.Add(value);
                    }
                }
            }

            if (values.Count != 1)
            {
                throw new InvalidOperationException(string.Format(
                    "{0}에서 환자 [{1}/{2}/{3}]의 고유 주민번호 암호문이 {4}개입니다. 정확히 1개일 때만 복구할 수 있습니다.",
                    sourceDescription, chartNo, patientName, juminPrefix, values.Count));
            }
            return values[0];
        }

        private string BuildJuminRestoreWhereClause(JuminRestoreTablePlan table)
        {
            return "CONVERT(nvarchar(30), " + QuoteSqlName(table.ChartColumn) + ") = @chrtNo" +
                   " AND LTRIM(RTRIM(ISNULL(CONVERT(nvarchar(100), " + QuoteSqlName(table.NameColumn) + "), N''))) = LTRIM(RTRIM(@patNm))" +
                   " AND LEFT(REPLACE(REPLACE(ISNULL(CONVERT(nvarchar(50), " + QuoteSqlName(table.JuminColumn) + "), N''), '-', ''), ' ', ''), 7) = @prefix" +
                   " AND ISNULL(CONVERT(nvarchar(max), JUMIN_ENCRYPT), N'') = @oldCipher";
        }

        private void AddJuminRestoreParameters(SqlCommand cmd, JuminEncryptionRestorePlan plan)
        {
            cmd.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = plan.ChartNo;
            cmd.Parameters.Add("@patNm", SqlDbType.NVarChar, 100).Value = plan.PatientName;
            cmd.Parameters.Add("@prefix", SqlDbType.NVarChar, 7).Value = plan.JuminPrefix;
            cmd.Parameters.Add("@oldCipher", SqlDbType.NVarChar, -1).Value = plan.OldCipher;
        }

        private int CountJuminRestoreRows(SqlConnection conn, SqlTransaction trans,
            JuminRestoreTablePlan table, JuminEncryptionRestorePlan plan, bool lockRows)
        {
            if (!TableExists(conn, table.TableName, trans)) return 0;

            string lockHint = lockRows ? " WITH (UPDLOCK, HOLDLOCK)" : " WITH (NOLOCK)";
            string sql = "SELECT COUNT(*) FROM dbo." + QuoteSqlName(table.TableName) + lockHint +
                         " WHERE " + BuildJuminRestoreWhereClause(table) + ";";
            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            {
                cmd.CommandTimeout = 120;
                AddJuminRestoreParameters(cmd, plan);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private int CountUnresolvedBackupPatientMismatches(SqlConnection conn, SqlTransaction trans,
            string chartNo, string backupDatabase, bool lockRows)
        {
            string currentLock = lockRows ? " WITH (UPDLOCK, HOLDLOCK)" : " WITH (NOLOCK)";
            string sql = @"
SELECT COUNT(*)
FROM dbo.TBSID040_03 r" + currentLock + @"
INNER JOIN " + QuoteSqlName(backupDatabase) + @".dbo.TBSID040_03 b WITH (NOLOCK)
    ON b.DRUG_SEQ = r.DRUG_SEQ
WHERE r.CHRTNO = @chrtNo
  AND
  (
       LTRIM(RTRIM(ISNULL(r.PAT_NM, N''))) <> LTRIM(RTRIM(ISNULL(b.PAT_NM, N'')))
    OR (
         NULLIF(LEFT(REPLACE(REPLACE(ISNULL(r.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7), N'') IS NOT NULL
     AND NULLIF(LEFT(REPLACE(REPLACE(ISNULL(b.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7), N'') IS NOT NULL
     AND LEFT(REPLACE(REPLACE(ISNULL(r.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7)
         <> LEFT(REPLACE(REPLACE(ISNULL(b.PAT_JUMIN_NO, N''), N'-', N''), N' ', N''), 7)
       )
  );";
            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
            {
                cmd.CommandTimeout = 120;
                cmd.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = chartNo;
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private JuminEncryptionRestorePlan BuildJuminEncryptionRestorePlan(SqlConnection conn, SqlTransaction trans,
            string chartNo, string patientName, string juminPrefix, string backupDatabase, bool lockRows)
        {
            if (string.IsNullOrWhiteSpace(chartNo) || string.IsNullOrWhiteSpace(patientName) ||
                !Regex.IsMatch(juminPrefix ?? "", @"^\d{7}$"))
            {
                throw new InvalidOperationException("선택 환자의 차트번호, 환자명 또는 주민번호 앞 7자리가 올바르지 않습니다.");
            }

            string currentLock = lockRows ? " WITH (UPDLOCK, HOLDLOCK)" : " WITH (NOLOCK)";
            string currentSql = @"
SELECT DISTINCT CONVERT(nvarchar(4000), JUMIN_ENCRYPT)
FROM dbo.TBSIT000_01" + currentLock + @"
WHERE CHRTNO = @chrtNo
  AND CUSACT = '1'
  AND LTRIM(RTRIM(ISNULL(PAT_NM, N''))) = LTRIM(RTRIM(@patNm))
  AND LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, ''), '-', ''), ' ', ''), 7) = @prefix;";

            string backupSql = @"
SELECT DISTINCT CONVERT(nvarchar(4000), JUMIN_ENCRYPT)
FROM " + QuoteSqlName(backupDatabase) + @".dbo.TBSIT000_01 WITH (NOLOCK)
WHERE CUSACT = '1'
  AND LTRIM(RTRIM(ISNULL(PAT_NM, N''))) = LTRIM(RTRIM(@patNm))
  AND LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, ''), '-', ''), ' ', ''), 7) = @prefix
  AND NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), JUMIN_ENCRYPT))), N'') IS NOT NULL;";

            JuminEncryptionRestorePlan plan = new JuminEncryptionRestorePlan
            {
                ChartNo = chartNo.Trim(),
                PatientName = patientName.Trim(),
                JuminPrefix = juminPrefix,
                BackupDatabase = backupDatabase
            };

            int unresolvedBackupMismatches = CountUnresolvedBackupPatientMismatches(
                conn, trans, plan.ChartNo, backupDatabase, lockRows);
            if (unresolvedBackupMismatches > 0)
            {
                throw new InvalidOperationException(string.Format(
                    "현재 차트 [{0}]에 백업 원환자 불일치 처방이 {1:N0}건 남아 있습니다. " +
                    "[백업 원환자와 불일치]에서 해당 처방을 먼저 분리/복구한 뒤 주민번호 암호문 검사를 다시 실행하십시오.",
                    plan.ChartNo, unresolvedBackupMismatches));
            }

            plan.OldCipher = ReadSingleJuminCipher(conn, trans, currentSql, plan.ChartNo,
                plan.PatientName, plan.JuminPrefix, "현재 환자 마스터", true);
            plan.BackupCipher = ReadSingleJuminCipher(conn, trans, backupSql, plan.ChartNo,
                plan.PatientName, plan.JuminPrefix, "읽기 전용 백업 환자 마스터", false);

            if (string.Equals(plan.OldCipher, plan.BackupCipher, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("현재 암호문과 백업 암호문이 이미 같습니다. 복구할 변경 사항이 없습니다.");
            }

            plan.Tables = CreateJuminRestoreTablePlans();
            foreach (JuminRestoreTablePlan table in plan.Tables)
            {
                table.RowCount = CountJuminRestoreRows(conn, trans, table, plan, lockRows);
            }

            JuminRestoreTablePlan master = plan.Tables.First(t => t.TableName == "TBSIT000_01");
            if (master.RowCount < 1)
            {
                throw new InvalidOperationException("현재 환자 마스터에 복구할 관련 이력 행이 없습니다.");
            }
            if (plan.TotalRows < 1)
            {
                throw new InvalidOperationException("암호문을 복구할 대상 행이 없습니다.");
            }
            return plan;
        }

        private void EnsureJuminRestoreAuditTable(SqlConnection conn, SqlTransaction trans)
        {
            using (SqlCommand cmd = new SqlCommand(@"
IF OBJECT_ID(N'dbo.PM_HELPER_JUMIN_ENCRYPT_RESTORE_AUDIT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PM_HELPER_JUMIN_ENCRYPT_RESTORE_AUDIT
    (
        AUDIT_ID BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        BACKUP_DTIME DATETIME NOT NULL CONSTRAINT DF_PMHELPER_JUMIN_AUDIT_DTIME DEFAULT GETDATE(),
        SESSION_ID UNIQUEIDENTIFIER NOT NULL,
        TABLE_NAME NVARCHAR(128) NOT NULL,
        ROW_KEY NVARCHAR(500) NULL,
        CHRTNO NVARCHAR(20) NOT NULL,
        PAT_NM NVARCHAR(100) NOT NULL,
        JUMIN_PREFIX NVARCHAR(7) NOT NULL,
        OLD_JUMIN_ENCRYPT NVARCHAR(MAX) NULL,
        NEW_JUMIN_ENCRYPT NVARCHAR(MAX) NULL,
        WINDOWS_USER NVARCHAR(200) NULL
    );
    CREATE INDEX IX_PMHELPER_JUMIN_AUDIT_SESSION
        ON dbo.PM_HELPER_JUMIN_ENCRYPT_RESTORE_AUDIT(SESSION_ID, TABLE_NAME);
END;", conn, trans))
            {
                cmd.CommandTimeout = 120;
                cmd.ExecuteNonQuery();
            }
        }

        private void VerifyJuminRestorePlanUnchanged(JuminEncryptionRestorePlan preview,
            JuminEncryptionRestorePlan lockedPlan)
        {
            if (!string.Equals(preview.ChartNo, lockedPlan.ChartNo, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(preview.PatientName, lockedPlan.PatientName, StringComparison.Ordinal) ||
                !string.Equals(preview.JuminPrefix, lockedPlan.JuminPrefix, StringComparison.Ordinal) ||
                !string.Equals(preview.OldCipher, lockedPlan.OldCipher, StringComparison.Ordinal) ||
                !string.Equals(preview.BackupCipher, lockedPlan.BackupCipher, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("미리보기 후 환자 식별값이 변경되었습니다. 아무것도 변경하지 않고 작업을 취소합니다.");
            }

            foreach (JuminRestoreTablePlan previewTable in preview.Tables)
            {
                JuminRestoreTablePlan lockedTable = lockedPlan.Tables.First(t => t.TableName == previewTable.TableName);
                if (previewTable.RowCount != lockedTable.RowCount)
                {
                    throw new InvalidOperationException(string.Format(
                        "미리보기 후 {0} 행 수가 {1}건에서 {2}건으로 변경되었습니다. 아무것도 변경하지 않고 작업을 취소합니다.",
                        previewTable.TableName, previewTable.RowCount, lockedTable.RowCount));
                }
            }
        }

        private int ApplyJuminEncryptionRestoreInTransaction(SqlConnection conn, SqlTransaction trans,
            JuminEncryptionRestorePlan preview, Guid sessionId)
        {
            JuminEncryptionRestorePlan plan = BuildJuminEncryptionRestorePlan(conn, trans,
                preview.ChartNo, preview.PatientName, preview.JuminPrefix,
                preview.BackupDatabase, true);
            VerifyJuminRestorePlanUnchanged(preview, plan);

            int totalUpdated = 0;
            foreach (JuminRestoreTablePlan table in plan.Tables)
            {
                if (table.RowCount == 0) continue;

                string whereClause = BuildJuminRestoreWhereClause(table);
                string auditSql = @"
INSERT INTO dbo.PM_HELPER_JUMIN_ENCRYPT_RESTORE_AUDIT
    (SESSION_ID, TABLE_NAME, ROW_KEY, CHRTNO, PAT_NM, JUMIN_PREFIX,
     OLD_JUMIN_ENCRYPT, NEW_JUMIN_ENCRYPT, WINDOWS_USER)
SELECT @sessionId, @tableName, " + table.RowKeyExpression + @", @chrtNo, @patNm, @prefix,
       CONVERT(nvarchar(max), JUMIN_ENCRYPT), @newCipher, SYSTEM_USER
FROM dbo." + QuoteSqlName(table.TableName) + @" WITH (UPDLOCK, HOLDLOCK)
WHERE " + whereClause + ";";
                int auditRows;
                using (SqlCommand cmdAudit = new SqlCommand(auditSql, conn, trans))
                {
                    cmdAudit.CommandTimeout = 120;
                    AddJuminRestoreParameters(cmdAudit, plan);
                    cmdAudit.Parameters.Add("@sessionId", SqlDbType.UniqueIdentifier).Value = sessionId;
                    cmdAudit.Parameters.Add("@tableName", SqlDbType.NVarChar, 128).Value = table.TableName;
                    cmdAudit.Parameters.Add("@newCipher", SqlDbType.NVarChar, -1).Value = plan.BackupCipher;
                    auditRows = cmdAudit.ExecuteNonQuery();
                }
                if (auditRows != table.RowCount)
                {
                    throw new InvalidOperationException(string.Format(
                        "{0} 원본 백업 건수({1})와 예상 건수({2})가 다릅니다.",
                        table.TableName, auditRows, table.RowCount));
                }

                string updateSql = "UPDATE dbo." + QuoteSqlName(table.TableName) +
                                   " SET JUMIN_ENCRYPT = @newCipher WHERE " + whereClause + ";";
                int updatedRows;
                using (SqlCommand cmdUpdate = new SqlCommand(updateSql, conn, trans))
                {
                    cmdUpdate.CommandTimeout = 120;
                    AddJuminRestoreParameters(cmdUpdate, plan);
                    cmdUpdate.Parameters.Add("@newCipher", SqlDbType.NVarChar, -1).Value = plan.BackupCipher;
                    updatedRows = cmdUpdate.ExecuteNonQuery();
                }
                if (updatedRows != table.RowCount)
                {
                    throw new InvalidOperationException(string.Format(
                        "{0} 변경 건수({1})와 예상 건수({2})가 다릅니다.",
                        table.TableName, updatedRows, table.RowCount));
                }
                totalUpdated += updatedRows;
            }

            foreach (JuminRestoreTablePlan table in plan.Tables)
            {
                if (CountJuminRestoreRows(conn, trans, table, plan, true) != 0)
                    throw new InvalidOperationException(table.TableName + "에 이전 암호문이 남아 있어 전체 작업을 취소합니다.");
            }

            if (totalUpdated != plan.TotalRows)
                throw new InvalidOperationException("전체 변경 건수 검증에 실패하여 작업을 취소합니다.");
            return totalUpdated;
        }

        private JuminEncryptionRestoreResult ExecuteJuminEncryptionRestoreBatch(
            List<JuminEncryptionRestorePlan> previews, bool commitChanges)
        {
            if (previews == null || previews.Count == 0)
                throw new InvalidOperationException("일괄 복구할 환자가 선택되지 않았습니다.");

            HashSet<string> identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (JuminEncryptionRestorePlan preview in previews)
            {
                string key = preview.ChartNo + "|" + preview.PatientName + "|" + preview.JuminPrefix;
                if (!identities.Add(key))
                    throw new InvalidOperationException("같은 환자가 일괄 복구 목록에 중복으로 포함되어 있습니다: " + preview.ChartNo);
            }

            using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
            {
                conn.Open();
                using (SqlTransaction trans = conn.BeginTransaction(IsolationLevel.Serializable))
                {
                    try
                    {
                        EnsureJuminRestoreAuditTable(conn, trans);
                        Guid sessionId = Guid.NewGuid();
                        int totalUpdated = 0;
                        foreach (JuminEncryptionRestorePlan preview in previews)
                            totalUpdated += ApplyJuminEncryptionRestoreInTransaction(conn, trans, preview, sessionId);

                        if (commitChanges) trans.Commit();
                        else trans.Rollback();
                        return new JuminEncryptionRestoreResult
                        {
                            SessionId = sessionId,
                            UpdatedRows = totalUpdated,
                            Committed = commitChanges
                        };
                    }
                    catch
                    {
                        try { trans.Rollback(); } catch { }
                        throw;
                    }
                }
            }
        }

        private JuminEncryptionRestoreResult ExecuteJuminEncryptionRestore(
            JuminEncryptionRestorePlan preview, bool commitChanges)
        {
            return ExecuteJuminEncryptionRestoreBatch(
                new List<JuminEncryptionRestorePlan> { preview }, commitChanges);
        }

        private string FormatJuminRestoreCounts(JuminEncryptionRestorePlan plan)
        {
            StringBuilder sb = new StringBuilder();
            foreach (JuminRestoreTablePlan table in plan.Tables)
            {
                sb.AppendFormat("- {0} ({1}): {2:N0}건\n", table.DisplayName, table.TableName, table.RowCount);
            }
            sb.AppendFormat("- 합계: {0:N0}건", plan.TotalRows);
            return sb.ToString();
        }

        private string FormatJuminBatchRestoreCounts(List<JuminEncryptionRestorePlan> plans)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>();
            Dictionary<string, string> names = new Dictionary<string, string>();
            foreach (JuminEncryptionRestorePlan plan in plans)
            {
                foreach (JuminRestoreTablePlan table in plan.Tables)
                {
                    if (!counts.ContainsKey(table.TableName)) counts[table.TableName] = 0;
                    counts[table.TableName] += table.RowCount;
                    names[table.TableName] = table.DisplayName;
                }
            }

            StringBuilder sb = new StringBuilder();
            foreach (JuminRestoreTablePlan table in CreateJuminRestoreTablePlans())
            {
                int count = counts.ContainsKey(table.TableName) ? counts[table.TableName] : 0;
                sb.AppendFormat("- {0} ({1}): {2:N0}건\n", names.ContainsKey(table.TableName) ? names[table.TableName] : table.DisplayName,
                    table.TableName, count);
            }
            sb.AppendFormat("- 합계: {0:N0}건", plans.Sum(p => p.TotalRows));
            return sb.ToString();
        }

        private void RestoreSelectedJuminEncryptionMismatches()
        {
            List<DataGridViewRow> selectedRows = GetCheckedJuminRestoreRows();
            if (selectedRows.Count == 0)
            {
                MessageBox.Show("왼쪽에서 복구 가능 환자를 고른 뒤 오른쪽의 [현재 환자 전체선택]을 누르십시오.\n암호문 복구는 같은 환자의 관련 행 전체를 함께 처리합니다.", "환자 전체선택 필요",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_chkDemoMode != null && _chkDemoMode.Checked)
            {
                MessageBox.Show("주민번호 암호문 복구는 실제 DB 모드에서만 실행할 수 있습니다.", "실제 DB 필요",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool backupReadOnly;
            string backupDatabase = FindAttachedPrescriptionBackupDatabase(out backupReadOnly);
            if (string.IsNullOrEmpty(backupDatabase) || !backupReadOnly)
            {
                MessageBox.Show(
                    "읽기 전용으로 연결된 PM_MAIN 백업 DB가 필요합니다.\n[백업 DB 연결]을 먼저 실행하고 상태가 '읽기 전용'인지 확인하십시오.",
                    "안전한 백업 DB 필요", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<JuminEncryptionRestorePlan> previews = new List<JuminEncryptionRestorePlan>();
            try
            {
                this.Cursor = Cursors.WaitCursor;
                using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                {
                    conn.Open();
                    foreach (DataGridViewRow row in selectedRows.OrderBy(r => Convert.ToString(r.Cells["차트번호"].Value)))
                    {
                        string chartNo = Convert.ToString(row.Cells["차트번호"].Value).Trim();
                        string patientName = Convert.ToString(row.Cells["현재환자명"].Value).Trim();
                        string digits = Regex.Replace(Convert.ToString(row.Cells["주민등록번호"].Value) ?? "", @"[^0-9]", "");
                        if (digits.Length < 7)
                            throw new InvalidOperationException(string.Format("[{0}/{1}] 주민번호 앞 7자리를 확인할 수 없습니다.", chartNo, patientName));
                        previews.Add(BuildJuminEncryptionRestorePlan(conn, null, chartNo, patientName,
                            digits.Substring(0, 7), backupDatabase, false));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("복구 미리보기 검증에 실패했습니다. DB는 변경되지 않았습니다.\n\n" + ex.Message,
                    "복구 중단", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }

            string previewMessage = string.Format(
                "현재 환자의 잘못된 주민번호 암호문을 읽기 전용 백업 값으로 복구합니다.\n\n" +
                "- 백업 DB: {0} (읽기 전용)\n\n변경 예정\n{1}\n\n" +
                "주민번호 표시값과 다른 진료 데이터는 변경하지 않습니다.\n" +
                "환자 마스터·처방·청구 등 같은 환자의 관련 암호문 전체가 하나의 트랜잭션으로 처리됩니다.\n" +
                "변경 전 암호문은 동일한 감사 세션 ID로 보관됩니다.",
                backupDatabase, FormatJuminBatchRestoreCounts(previews));
            if (MessageBox.Show(previewMessage, "주민번호 암호문 복구 미리보기",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }
            if (MessageBox.Show(
                "PMPLUS20에서 현재 환자 화면을 닫고 다른 사용자가 수정 중이 아닌지 확인했습니까?\n\n계속하면 위 합계의 암호문을 실제 DB에 반영합니다.",
                "암호문 복구 최종 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;
                JuminEncryptionRestoreResult result = ExecuteJuminEncryptionRestoreBatch(previews, true);
                MessageBox.Show(string.Format(
                    "주민번호 암호문 복구가 완료되었습니다.\n\n- 환자: {0:N0}명\n- 변경: {1:N0}건\n- 감사 세션: {2}\n\n검사 목록을 다시 조회합니다.",
                    previews.Count, result.UpdatedRows, result.SessionId),
                    "암호문 복구 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ShowToast(string.Format("환자 {0:N0}명, 암호문 {1:N0}건 복구 완료",
                    previews.Count, result.UpdatedRows), ColorEmerald);
                BtnLogMismatchScan_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "암호문 복구에 실패하여 현재 환자의 전체 트랜잭션을 취소했습니다. DB에는 일부 변경이 남지 않습니다.\n\n" + ex.Message,
                    "암호문 복구 실패 및 전체 취소", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void BtnLogRestoreSeparate_Click(object sender, EventArgs e)
        {
            if (_cmbLogMismatchFilter != null && _cmbLogMismatchFilter.SelectedIndex == 7)
            {
                RestoreSelectedJuminEncryptionMismatches();
                return;
            }

            if (_dgvLogMismatchDetail == null || _dgvLogMismatchDetail.Rows.Count == 0)
            {
                MessageBox.Show("분리/복구할 조제 내역이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<string> selectedDrugSeqs = new List<string>();
            string targetPatientName = "";
            HashSet<string> selectedPatientNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> selectedOriginalCharts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> selectedLogJumins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> selectedSourceCharts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool containsNonRestorableRow = false;

            foreach (DataGridViewRow row in _dgvLogMismatchDetail.Rows)
            {
                if (row.Cells["선택"] != null && Convert.ToBoolean(row.Cells["선택"].Value))
                {
                    if (row.Cells["복구가능"] == null || !Convert.ToBoolean(row.Cells["복구가능"].Value))
                    {
                        containsNonRestorableRow = true;
                        continue;
                    }
                    string dseq = Convert.ToString(row.Cells["조제번호"].Value);
                    if (!string.IsNullOrEmpty(dseq))
                    {
                        if (!selectedDrugSeqs.Contains(dseq)) selectedDrugSeqs.Add(dseq);
                        string rowSourceChart = Convert.ToString(row.Cells["원장차트"].Value).Trim();
                        if (!string.IsNullOrEmpty(rowSourceChart)) selectedSourceCharts.Add(rowSourceChart);
                        if (string.IsNullOrEmpty(targetPatientName))
                        {
                            targetPatientName = Convert.ToString(row.Cells["복구환자명"].Value);
                        }
                        string rowPatientName = Convert.ToString(row.Cells["복구환자명"].Value).Trim();
                        if (!string.IsNullOrEmpty(rowPatientName))
                        {
                            selectedPatientNames.Add(rowPatientName);
                        }
                        string originalChart = Convert.ToString(row.Cells["복구원차트"].Value).Trim();
                        if (!string.IsNullOrEmpty(originalChart)) selectedOriginalCharts.Add(originalChart);
                        string logJumin = Convert.ToString(row.Cells["복구주민번호"].Value).Trim();
                        if (!string.IsNullOrEmpty(logJumin)) selectedLogJumins.Add(logJumin);
                    }
                }
            }

            if (selectedDrugSeqs.Count == 0)
            {
                MessageBox.Show("분리/복구할 처방전을 최소 1개 이상 체크(선택)하십시오.", "선택 필요", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (containsNonRestorableRow)
            {
                MessageBox.Show("복구 근거가 부족한 행이 선택되어 있습니다. '복구가능'이 체크된 행만 선택하십시오.", "복구 불가 행 선택", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (selectedSourceCharts.Count != 1)
            {
                MessageBox.Show("선택된 처방의 현재 원장차트가 서로 다릅니다. 한 원장차트씩 처리하십시오.", "원장차트 혼합", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sourceChrtNo = selectedSourceCharts.First();

            if (selectedPatientNames.Count > 1)
            {
                MessageBox.Show(
                    "서로 다른 원본 환자의 처방이 함께 선택되어 있습니다.\n같은 환자의 처방만 선택하여 한 차트씩 복구하십시오.",
                    "환자 혼합 선택",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (selectedOriginalCharts.Count != 1)
            {
                MessageBox.Show(
                    "선택된 처방의 복구 원차트번호가 없거나 서로 다릅니다.\n'복구원차트'가 같은 행만 선택하여 한 차트씩 복구하십시오.",
                    "원차트 단위 선택 필요",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (selectedLogJumins.Count > 1)
            {
                MessageBox.Show(
                    "선택된 처방의 원본 주민번호가 서로 다릅니다. 환자 오선택 가능성이 있어 복구를 중단합니다.",
                    "주민번호 불일치",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string selectedLogJumin = selectedLogJumins.Count == 1 ? selectedLogJumins.First() : "";
            string logJuminDigits = Regex.Replace(selectedLogJumin, @"[^0-9]", "");
            if (logJuminDigits.Length < 7)
            {
                MessageBox.Show(
                    "백업/로그 주민번호 앞 7자리를 확인할 수 없어 동일 환자 여부를 안전하게 검증할 수 없습니다.\n이 처방은 자동 복구하지 말고 CSV 결과를 확인하십시오.",
                    "환자 식별 근거 부족",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            string logJuminPrefix = logJuminDigits.Substring(0, 7);

            string newChrtNo = _txtLogRestoreNewChrtNo.Text.Trim();
            string logOriginalChart = selectedOriginalCharts.First();
            if (string.IsNullOrEmpty(newChrtNo))
            {
                MessageBox.Show("분리하여 할당할 대상 환자의 차트번호를 입력하십시오.", "차트번호 필요", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtLogRestoreNewChrtNo.Focus();
                return;
            }

            if (!string.Equals(newChrtNo, logOriginalChart, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    string.Format("입력한 차트번호 [{0}]가 선택 처방의 복구 원차트 [{1}]와 다릅니다.\n백업/로그 근거의 원차트번호를 확인하십시오.", newChrtNo, logOriginalChart),
                    "원차트번호 불일치",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!Regex.IsMatch(newChrtNo, @"^\d{10}$"))
            {
                MessageBox.Show("차트번호는 0을 포함한 10자리 숫자여야 합니다.", "차트번호 형식 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(targetPatientName))
            {
                targetPatientName = "복구환자";
            }

            string confirmMsg = string.Format(
                "선택한 {0}건의 조제 내역을 아래의 대상 환자 차트로 분리/복구하시겠습니까?\n\n" +
                "- 대상 차트번호: {1}\n" +
                "- 복구 환자명: {2}\n" +
                "- 변경 적용 테이블: TBSID040_03 (처방·조제 마스터)\n" +
                "- 상세 테이블 TBSID040_04/05는 DRUG_SEQ로 연결되므로 변경하지 않음\n\n" +
                "※ 변경 전 PM_HELPER_CHART_RESTORE_BACKUP 테이블에 원본 값을 영구 백업합니다.\n" +
                "※ 대상 차트가 없으면 현재 고객 원장을 먼저 확인하고, 없을 경우 읽기 전용 백업 DB의 해당 원차트 고객 이력을 복원합니다.\n" +
                "※ 환자명·주민번호 앞 7자리·암호문·활성행이 유일하게 검증되지 않으면 작업을 중단합니다.",
                selectedDrugSeqs.Count, newChrtNo, targetPatientName);

            if (MessageBox.Show(confirmMsg, "처방 분리/복구 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (_chkDemoMode.Checked)
                {
                    ShowToast(string.Format("[데모] {0}건의 처방전이 차트 [{1}] {2}님으로 분리 복구되었습니다.", selectedDrugSeqs.Count, newChrtNo, targetPatientName), ColorEmerald);
                    MessageBox.Show(string.Format("성공적으로 {0}건의 처방이 차트 [{1}] {2}님으로 분리 복구되었습니다.\n(데모 모드 시뮬레이션)", selectedDrugSeqs.Count, newChrtNo, targetPatientName), "복구 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    BtnLogMismatchScan_Click(null, null);
                }
                else
                {
                    using (SqlConnection conn = new SqlConnection(BuildConnectionString(false)))
                    {
                        conn.Open();
                        string customerActiveColumn = FindFirstColumn(conn, "TBSIT000_01", new string[] { "CUSACT", "CUS_ACT" });
                        if (string.IsNullOrEmpty(customerActiveColumn))
                        {
                            throw new InvalidOperationException("TBSIT000_01에서 고객 활성 컬럼(CUSACT/CUS_ACT)을 찾지 못했습니다.");
                        }

                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                Guid restoreSessionId = Guid.NewGuid();
                                string customerRestoreSource = "현재 고객 원장";
                                int restoredCustomerRows = 0;

                                // 1. Resolve verified patient identity from the target chart or
                                // exactly one existing customer with the same name/JUMIN prefix.
                                // If it no longer exists in the current DB, restore the complete
                                // customer history for the exact original chart from the read-only backup DB.
                                string verifiedJumin = "";
                                string verifiedEncrypt = "";
                                string verifiedFamNm = targetPatientName;
                                string existingTargetName = "";
                                using (SqlCommand cmdTarget = new SqlCommand(@"
SELECT TOP (1) PAT_NM, JUMIN_NO, JUMIN_ENCRYPT, FAM_NM
FROM dbo.TBSIT000_01 WITH (UPDLOCK, HOLDLOCK)
WHERE CHRTNO = @chrtNo
ORDER BY PAT_SEQ;", conn, trans))
                                {
                                    cmdTarget.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                    using (SqlDataReader reader = cmdTarget.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            existingTargetName = Convert.ToString(reader["PAT_NM"]).Trim();
                                            verifiedJumin = Convert.ToString(reader["JUMIN_NO"]);
                                            verifiedEncrypt = Convert.ToString(reader["JUMIN_ENCRYPT"]);
                                            verifiedFamNm = Convert.ToString(reader["FAM_NM"]);
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(existingTargetName) &&
                                    !string.Equals(existingTargetName, targetPatientName, StringComparison.OrdinalIgnoreCase))
                                {
                                    throw new InvalidOperationException(string.Format(
                                        "복구 원차트 [{0}]가 이미 다른 환자 [{1}]의 고객번호로 사용 중입니다.",
                                        newChrtNo,
                                        existingTargetName));
                                }

                                if (string.IsNullOrEmpty(existingTargetName))
                                {
                                    DataTable sourceCustomers = new DataTable();
                                    string sourceCustomerSql = @"
SELECT JUMIN_NO, JUMIN_ENCRYPT, MAX(NULLIF(FAM_NM, '')) AS FAM_NM
FROM dbo.TBSIT000_01 WITH (UPDLOCK, HOLDLOCK)
WHERE PAT_NM = @patNm
  AND REPLACE(REPLACE(ISNULL(JUMIN_NO, ''), '-', ''), ' ', '') LIKE @juminPrefix + '%'
GROUP BY JUMIN_NO, JUMIN_ENCRYPT;";
                                    using (SqlCommand cmdSource = new SqlCommand(sourceCustomerSql, conn, trans))
                                    using (SqlDataAdapter sourceAdapter = new SqlDataAdapter(cmdSource))
                                    {
                                        cmdSource.Parameters.Add("@patNm", SqlDbType.NVarChar, 50).Value = targetPatientName;
                                        cmdSource.Parameters.Add("@juminPrefix", SqlDbType.NVarChar, 20).Value = logJuminPrefix;
                                        sourceAdapter.Fill(sourceCustomers);
                                    }

                                    if (sourceCustomers.Rows.Count > 1)
                                    {
                                        throw new InvalidOperationException(string.Format(
                                            "원본 환자 [{0}/{1}]와 일치하는 고유 고객 식별값이 {2}개입니다. 정확히 1개일 때만 원차트를 자동 생성할 수 있습니다.",
                                            targetPatientName,
                                            selectedLogJumin,
                                            sourceCustomers.Rows.Count));
                                    }

                                    if (sourceCustomers.Rows.Count == 1)
                                    {
                                        DataRow sourceCustomer = sourceCustomers.Rows[0];
                                        verifiedJumin = Convert.ToString(sourceCustomer["JUMIN_NO"]);
                                        verifiedEncrypt = Convert.ToString(sourceCustomer["JUMIN_ENCRYPT"]);
                                        verifiedFamNm = Convert.ToString(sourceCustomer["FAM_NM"]);

                                        if (string.IsNullOrEmpty(verifiedJumin) || string.IsNullOrEmpty(verifiedEncrypt))
                                        {
                                            throw new InvalidOperationException("기존 고객의 주민번호 또는 암호화 식별값이 없어 안전하게 원차트를 생성할 수 없습니다.");
                                        }

                                        string insertCustSql = @"
INSERT INTO dbo.TBSIT000_01
    (CHRTNO, PAT_SEQ, PAT_NM, JUMIN_NO, JUMIN_ENCRYPT, FAM_NM, " + QuoteSqlName(customerActiveColumn) + @", PROC_DTIME)
VALUES
    (@chrtNo, 1, @patNm, @jumin, @juminEncrypt, @famNm, '1', CONVERT(VARCHAR(8), GETDATE(), 112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(), 108), ':', ''));";
                                        using (SqlCommand cmdInsertCust = new SqlCommand(insertCustSql, conn, trans))
                                        {
                                            cmdInsertCust.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                            cmdInsertCust.Parameters.Add("@patNm", SqlDbType.NVarChar, 50).Value = targetPatientName;
                                            cmdInsertCust.Parameters.Add("@jumin", SqlDbType.NVarChar, 30).Value = verifiedJumin;
                                            cmdInsertCust.Parameters.Add("@juminEncrypt", SqlDbType.NVarChar, -1).Value = verifiedEncrypt;
                                            cmdInsertCust.Parameters.Add("@famNm", SqlDbType.NVarChar, 50).Value = string.IsNullOrEmpty(verifiedFamNm) ? targetPatientName : verifiedFamNm;
                                            restoredCustomerRows = cmdInsertCust.ExecuteNonQuery();
                                        }
                                    }
                                    else
                                    {
                                        bool backupReadOnly;
                                        string backupDatabaseName = FindAttachedPrescriptionBackupDatabase(out backupReadOnly);
                                        if (string.IsNullOrEmpty(backupDatabaseName) || !backupReadOnly)
                                        {
                                            throw new InvalidOperationException(
                                                "현재 고객 원장에 원본 환자가 없습니다. 원차트를 복구하려면 [백업 DB 연결]로 읽기 전용 PM_MAIN 백업 DB를 먼저 연결하십시오.");
                                        }

                                        string backupIdentitySql = @"
SELECT COUNT(*) AS TOTAL_ROWS,
       SUM(CASE WHEN ISNULL(" + QuoteSqlName(customerActiveColumn) + @", N'') = N'1' THEN 1 ELSE 0 END) AS ACTIVE_ROWS,
       COUNT(DISTINCT LTRIM(RTRIM(ISNULL(PAT_NM, N'')))) AS NAME_COUNT,
       MAX(LTRIM(RTRIM(ISNULL(PAT_NM, N'')))) AS PAT_NM,
       COUNT(DISTINCT LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, N''), N'-', N''), N' ', N''), 7)) AS PREFIX_COUNT,
       MAX(LEFT(REPLACE(REPLACE(ISNULL(JUMIN_NO, N''), N'-', N''), N' ', N''), 7)) AS JUMIN_PREFIX,
       COUNT(DISTINCT NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), JUMIN_ENCRYPT))), N'')) AS CIPHER_COUNT,
       SUM(CASE WHEN NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(4000), JUMIN_ENCRYPT))), N'') IS NULL THEN 1 ELSE 0 END) AS EMPTY_CIPHER_ROWS,
       MAX(CONVERT(nvarchar(4000), JUMIN_ENCRYPT)) AS JUMIN_ENCRYPT,
       MAX(CASE WHEN ISNULL(" + QuoteSqlName(customerActiveColumn) + @", N'') = N'1' THEN JUMIN_NO END) AS ACTIVE_JUMIN_NO,
       MAX(CASE WHEN ISNULL(" + QuoteSqlName(customerActiveColumn) + @", N'') = N'1' THEN FAM_NM END) AS ACTIVE_FAM_NM
FROM " + QuoteSqlName(backupDatabaseName) + @".dbo.TBSIT000_01 WITH (NOLOCK)
WHERE CHRTNO = @chrtNo;";

                                        int backupTotalRows;
                                        int backupActiveRows;
                                        int backupNameCount;
                                        int backupPrefixCount;
                                        int backupCipherCount;
                                        int backupEmptyCipherRows;
                                        string backupPatientName;
                                        string backupJuminPrefix;
                                        using (SqlCommand cmdBackupIdentity = new SqlCommand(backupIdentitySql, conn, trans))
                                        {
                                            cmdBackupIdentity.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                            using (SqlDataReader reader = cmdBackupIdentity.ExecuteReader())
                                            {
                                                if (!reader.Read())
                                                {
                                                    throw new InvalidOperationException("백업 고객 원장을 읽지 못했습니다.");
                                                }
                                                backupTotalRows = reader.IsDBNull(reader.GetOrdinal("TOTAL_ROWS")) ? 0 : Convert.ToInt32(reader["TOTAL_ROWS"]);
                                                backupActiveRows = reader.IsDBNull(reader.GetOrdinal("ACTIVE_ROWS")) ? 0 : Convert.ToInt32(reader["ACTIVE_ROWS"]);
                                                backupNameCount = reader.IsDBNull(reader.GetOrdinal("NAME_COUNT")) ? 0 : Convert.ToInt32(reader["NAME_COUNT"]);
                                                backupPatientName = Convert.ToString(reader["PAT_NM"]).Trim();
                                                backupPrefixCount = reader.IsDBNull(reader.GetOrdinal("PREFIX_COUNT")) ? 0 : Convert.ToInt32(reader["PREFIX_COUNT"]);
                                                backupJuminPrefix = Convert.ToString(reader["JUMIN_PREFIX"]).Trim();
                                                backupCipherCount = reader.IsDBNull(reader.GetOrdinal("CIPHER_COUNT")) ? 0 : Convert.ToInt32(reader["CIPHER_COUNT"]);
                                                backupEmptyCipherRows = reader.IsDBNull(reader.GetOrdinal("EMPTY_CIPHER_ROWS")) ? 0 : Convert.ToInt32(reader["EMPTY_CIPHER_ROWS"]);
                                                verifiedEncrypt = Convert.ToString(reader["JUMIN_ENCRYPT"]);
                                                verifiedJumin = Convert.ToString(reader["ACTIVE_JUMIN_NO"]);
                                                verifiedFamNm = Convert.ToString(reader["ACTIVE_FAM_NM"]);
                                            }
                                        }

                                        if (backupTotalRows < 1)
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "읽기 전용 백업 DB [{0}]에도 원차트 [{1}]의 고객 원장이 없습니다.", backupDatabaseName, newChrtNo));
                                        }
                                        if (backupNameCount != 1 ||
                                            !string.Equals(backupPatientName, targetPatientName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "백업 원차트 [{0}]의 환자명이 선택 환자 [{1}]로 유일하게 확인되지 않습니다.", newChrtNo, targetPatientName));
                                        }
                                        if (backupPrefixCount != 1 || !string.Equals(backupJuminPrefix, logJuminPrefix, StringComparison.Ordinal))
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "백업 원차트 [{0}]의 주민번호 앞 7자리가 처방 근거 [{1}]와 유일하게 일치하지 않습니다.", newChrtNo, logJuminPrefix));
                                        }
                                        if (backupCipherCount != 1 || backupEmptyCipherRows != 0 || string.IsNullOrEmpty(verifiedEncrypt))
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "백업 원차트 [{0}]의 암호화 주민 식별값이 하나로 확정되지 않습니다.", newChrtNo));
                                        }
                                        if (backupActiveRows != 1 || string.IsNullOrEmpty(verifiedJumin))
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "백업 원차트 [{0}]의 활성 고객 행이 {1}개입니다. 정확히 1개일 때만 복구할 수 있습니다.", newChrtNo, backupActiveRows));
                                        }

                                        List<string> cloneColumns = GetBackupCustomerCloneColumns(conn, trans, backupDatabaseName);
                                        string cloneColumnList = string.Join(", ", cloneColumns.Select(QuoteSqlName).ToArray());
                                        string cloneSql = @"
INSERT INTO dbo.TBSIT000_01 (" + cloneColumnList + @")
SELECT " + cloneColumnList + @"
FROM " + QuoteSqlName(backupDatabaseName) + @".dbo.TBSIT000_01 WITH (NOLOCK)
WHERE CHRTNO = @chrtNo;";
                                        using (SqlCommand cmdCloneCustomer = new SqlCommand(cloneSql, conn, trans))
                                        {
                                            cmdCloneCustomer.CommandTimeout = 120;
                                            cmdCloneCustomer.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                            restoredCustomerRows = cmdCloneCustomer.ExecuteNonQuery();
                                        }
                                        if (restoredCustomerRows != backupTotalRows)
                                        {
                                            throw new InvalidOperationException(string.Format(
                                                "백업 고객 이력은 {0}행이지만 실제 복원된 고객 이력은 {1}행입니다. 전체 작업을 취소합니다.",
                                                backupTotalRows, restoredCustomerRows));
                                        }

                                        using (SqlCommand cmdCustomerAudit = new SqlCommand(@"
IF OBJECT_ID(N'dbo.PM_HELPER_CUSTOMER_RESTORE_AUDIT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PM_HELPER_CUSTOMER_RESTORE_AUDIT
    (
        AUDIT_ID BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        RESTORE_DTIME DATETIME NOT NULL DEFAULT GETDATE(),
        SESSION_ID UNIQUEIDENTIFIER NOT NULL,
        SOURCE_DATABASE NVARCHAR(128) NOT NULL,
        CHRTNO NVARCHAR(20) NOT NULL,
        PAT_NM NVARCHAR(100) NOT NULL,
        JUMIN_PREFIX NVARCHAR(7) NOT NULL,
        INSERTED_ROWS INT NOT NULL,
        WINDOWS_USER NVARCHAR(200) NULL
    );
END;
INSERT INTO dbo.PM_HELPER_CUSTOMER_RESTORE_AUDIT
    (SESSION_ID, SOURCE_DATABASE, CHRTNO, PAT_NM, JUMIN_PREFIX, INSERTED_ROWS, WINDOWS_USER)
VALUES
    (@sessionId, @sourceDatabase, @chrtNo, @patNm, @juminPrefix, @insertedRows, SYSTEM_USER);", conn, trans))
                                        {
                                            cmdCustomerAudit.Parameters.Add("@sessionId", SqlDbType.UniqueIdentifier).Value = restoreSessionId;
                                            cmdCustomerAudit.Parameters.Add("@sourceDatabase", SqlDbType.NVarChar, 128).Value = backupDatabaseName;
                                            cmdCustomerAudit.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                            cmdCustomerAudit.Parameters.Add("@patNm", SqlDbType.NVarChar, 100).Value = targetPatientName;
                                            cmdCustomerAudit.Parameters.Add("@juminPrefix", SqlDbType.NVarChar, 7).Value = logJuminPrefix;
                                            cmdCustomerAudit.Parameters.Add("@insertedRows", SqlDbType.Int).Value = restoredCustomerRows;
                                            cmdCustomerAudit.ExecuteNonQuery();
                                        }
                                        customerRestoreSource = "읽기 전용 백업 DB " + backupDatabaseName;
                                    }
                                }

                                string verifiedDigits = Regex.Replace(verifiedJumin, @"[^0-9]", "");
                                if (string.IsNullOrEmpty(verifiedEncrypt))
                                {
                                    throw new InvalidOperationException("검증된 고객 마스터의 암호화 주민 식별값이 없어 자동 복구할 수 없습니다.");
                                }
                                if (verifiedDigits.Length < 7 || !verifiedDigits.StartsWith(logJuminPrefix))
                                {
                                    throw new InvalidOperationException("고객 마스터와 처방로그의 주민번호 앞 7자리가 일치하지 않아 복구를 중단합니다.");
                                }

                                // 2. Build a parameterized prescription set.
                                List<string> drugSeqParameters = new List<string>();
                                for (int i = 0; i < selectedDrugSeqs.Count; i++)
                                {
                                    drugSeqParameters.Add("@drugSeq" + i.ToString());
                                }
                                string inClause = string.Join(",", drugSeqParameters.ToArray());

                                // 3. Persist original values before any prescription change.
                                using (SqlCommand cmdBackupTable = new SqlCommand(@"
IF OBJECT_ID(N'dbo.PM_HELPER_CHART_RESTORE_BACKUP', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PM_HELPER_CHART_RESTORE_BACKUP
    (
        BACKUP_ID BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        BACKUP_DTIME DATETIME NOT NULL DEFAULT GETDATE(),
        SESSION_ID UNIQUEIDENTIFIER NOT NULL,
        DRUG_SEQ NVARCHAR(50) NOT NULL,
        OLD_CHRTNO NVARCHAR(20) NULL,
        OLD_PAT_NM NVARCHAR(100) NULL,
        OLD_PAT_JUMIN_NO NVARCHAR(50) NULL,
        OLD_JUMIN_ENCRYPT NVARCHAR(MAX) NULL,
        NEW_CHRTNO NVARCHAR(20) NOT NULL,
        NEW_PAT_NM NVARCHAR(100) NOT NULL,
        LOG_ORIGINAL_CHRTNO NVARCHAR(20) NOT NULL,
        WINDOWS_USER NVARCHAR(200) NULL
    );
END;", conn, trans))
                                {
                                    cmdBackupTable.ExecuteNonQuery();
                                }

                                string backupSql = string.Format(@"
INSERT INTO dbo.PM_HELPER_CHART_RESTORE_BACKUP
    (SESSION_ID, DRUG_SEQ, OLD_CHRTNO, OLD_PAT_NM, OLD_PAT_JUMIN_NO, OLD_JUMIN_ENCRYPT,
     NEW_CHRTNO, NEW_PAT_NM, LOG_ORIGINAL_CHRTNO, WINDOWS_USER)
SELECT @sessionId, DRUG_SEQ, CHRTNO, PAT_NM, PAT_JUMIN_NO, JUMIN_ENCRYPT,
       @chrtNo, @patNm, @logOriginalChart, SYSTEM_USER
FROM dbo.TBSID040_03 WITH (UPDLOCK, HOLDLOCK)
WHERE CHRTNO = @sourceChrtNo AND DRUG_SEQ IN ({0});", inClause);
                                using (SqlCommand cmdBackup = new SqlCommand(backupSql, conn, trans))
                                {
                                    cmdBackup.Parameters.Add("@sessionId", SqlDbType.UniqueIdentifier).Value = restoreSessionId;
                                    cmdBackup.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                    cmdBackup.Parameters.Add("@patNm", SqlDbType.NVarChar, 50).Value = targetPatientName;
                                    cmdBackup.Parameters.Add("@logOriginalChart", SqlDbType.NVarChar, 20).Value = logOriginalChart;
                                    cmdBackup.Parameters.Add("@sourceChrtNo", SqlDbType.NVarChar, 20).Value = sourceChrtNo;
                                    for (int i = 0; i < selectedDrugSeqs.Count; i++)
                                    {
                                        cmdBackup.Parameters.Add("@drugSeq" + i.ToString(), SqlDbType.NVarChar, 30).Value = selectedDrugSeqs[i];
                                    }
                                    int backupRows = cmdBackup.ExecuteNonQuery();
                                    if (backupRows != selectedDrugSeqs.Count)
                                    {
                                        throw new InvalidOperationException(string.Format(
                                            "선택한 처방은 {0}건이지만 백업된 처방은 {1}건입니다. 원장이 변경되었을 수 있어 작업을 취소합니다.",
                                            selectedDrugSeqs.Count,
                                            backupRows));
                                    }
                                }

                                // 4. Restore the master identity. TBSID040_04/05 remain linked by DRUG_SEQ.
                                string update03 = string.Format(@"
UPDATE dbo.TBSID040_03
SET CHRTNO = @chrtNo,
    PAT_NM = @patNm,
    PAT_JUMIN_NO = @patJumin,
    JUMIN_ENCRYPT = @juminEncrypt
WHERE CHRTNO = @sourceChrtNo AND DRUG_SEQ IN ({0});", inClause);
                                using (SqlCommand cmd03 = new SqlCommand(update03, conn, trans))
                                {
                                    cmd03.Parameters.Add("@chrtNo", SqlDbType.NVarChar, 20).Value = newChrtNo;
                                    cmd03.Parameters.Add("@patNm", SqlDbType.NVarChar, 50).Value = targetPatientName;
                                    cmd03.Parameters.Add("@patJumin", SqlDbType.NVarChar, 50).Value = verifiedJumin;
                                    cmd03.Parameters.Add("@juminEncrypt", SqlDbType.NVarChar, -1).Value = verifiedEncrypt;
                                    cmd03.Parameters.Add("@sourceChrtNo", SqlDbType.NVarChar, 20).Value = sourceChrtNo;
                                    for (int i = 0; i < selectedDrugSeqs.Count; i++)
                                    {
                                        cmd03.Parameters.Add("@drugSeq" + i.ToString(), SqlDbType.NVarChar, 30).Value = selectedDrugSeqs[i];
                                    }
                                    int updatedRows = cmd03.ExecuteNonQuery();
                                    if (updatedRows != selectedDrugSeqs.Count)
                                    {
                                        throw new InvalidOperationException(string.Format(
                                            "선택한 처방은 {0}건이지만 실제 변경된 처방은 {1}건입니다. 안전을 위해 작업을 취소합니다.",
                                            selectedDrugSeqs.Count,
                                            updatedRows));
                                    }
                                }

                                trans.Commit();

                                ShowToast(string.Format("{0}건의 처방이 차트 [{1}] {2}님으로 복구 완료되었습니다.", selectedDrugSeqs.Count, newChrtNo, targetPatientName), ColorEmerald);
                                MessageBox.Show(string.Format(
                                    "처방 분리/복구가 완료되었습니다.\n\n- 복구 건수: {0}건\n- 대상 차트번호: {1}\n- 대상 환자명: {2}\n- 고객 원장 복원: {3}행 ({4})\n- 백업 세션: {5}\n\n원본 값은 PM_HELPER_CHART_RESTORE_BACKUP에 보관되었습니다.",
                                    selectedDrugSeqs.Count,
                                    newChrtNo,
                                    targetPatientName,
                                    restoredCustomerRows,
                                    customerRestoreSource,
                                    restoreSessionId), "복구 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                BtnLogMismatchScan_Click(null, null);
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("처방 분리/복구 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void BtnLogMismatchExport_Click(object sender, EventArgs e)
        {
            if (_dgvLogMismatchSummary == null || _dgvLogMismatchSummary.Rows.Count == 0) return;

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 파일 (*.csv)|*.csv";
                dialog.FileName = string.Format("{0:yyyyMMdd}_로그청구불일치검사결과.csv", DateTime.Today);
                dialog.Title = "무결성 검사 결과 저장";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    using (StreamWriter writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true)))
                    {
                        List<DataGridViewColumn> visibleColumns = _dgvLogMismatchSummary.Columns
                            .Cast<DataGridViewColumn>()
                            .Where(c => c.Visible)
                            .OrderBy(c => c.DisplayIndex)
                            .ToList();

                        writer.WriteLine(string.Join(",", visibleColumns.Select(c => EscapeClaimCsvValue(c.HeaderText)).ToArray()));
                        foreach (DataGridViewRow row in _dgvLogMismatchSummary.Rows)
                        {
                            if (row.IsNewRow) continue;
                            writer.WriteLine(string.Join(",", visibleColumns.Select(c => EscapeClaimCsvValue(Convert.ToString(row.Cells[c.Index].Value))).ToArray()));
                        }
                    }

                    ShowToast("로그/청구 불일치 검사 결과를 CSV로 저장했습니다.", ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("CSV 저장 중 오류가 발생했습니다.\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Initialize Prescription Delete Tab Layout
        private void InitializePrescriptionDeleteTab()
        {
            // _tabPrescriptionDelete는 이미 _subTabDispenseCustomer에 추가되었습니다.

            // Top 검색 패널
            Panel pnlRxSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            _tabPrescriptionDelete.Controls.Add(pnlRxSearch);

            Label lblRxSearchName = new Label { Text = "환자명", Location = new Point(15, 20), Size = new Size(50, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtRxDelSearchName = new TextBox { Location = new Point(70, 17), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            Label lblRxSearchJumin = new Label { Text = "주민번호", Location = new Point(210, 20), Size = new Size(60, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtRxDelSearchJumin = new TextBox { Location = new Point(280, 17), Size = new Size(150, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            
            _btnRxDelSearch = new Button
            {
                Text = "🔍 처방 검색",
                Location = new Point(450, 14),
                Size = new Size(120, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnRxDelSearch.FlatAppearance.BorderSize = 0;
            _btnRxDelSearch.Click += BtnRxDelSearch_Click;

            pnlRxSearch.Controls.Add(lblRxSearchName);
            pnlRxSearch.Controls.Add(_txtRxDelSearchName);
            pnlRxSearch.Controls.Add(lblRxSearchJumin);
            pnlRxSearch.Controls.Add(_txtRxDelSearchJumin);
            pnlRxSearch.Controls.Add(_btnRxDelSearch);

            // SplitContainer
            _splitRx = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = _distRx,
                BackColor = ColorBorder
            };
            _tabPrescriptionDelete.Controls.Add(_splitRx);
            _splitRx.BringToFront();
            _splitRx.SplitterMoved += (s, e) => { _distRx = _splitRx.SplitterDistance; SaveConfig(); };
            _splitRx.Resize += (s, e) => NormalizeRightPanelSplit(_splitRx, ref _distRx, 340, 360);
            NormalizeRightPanelSplit(_splitRx, ref _distRx, 340, 360);

            _dgvRxDeleteList = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvRxDeleteList.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvRxDeleteList.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvRxDeleteList.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvRxDeleteList.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvRxDeleteList.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvRxDeleteList.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvRxDeleteList.DefaultCellStyle.SelectionForeColor = Color.White;
            _splitRx.Panel1.Controls.Add(_dgvRxDeleteList); // Panel1(왼쪽)에 목록 배치

            // Right Panel (삭제 안내 및 실행) -> Panel2에 배치
            Panel pnlRxForm = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(20)
            };
            _splitRx.Panel2.Controls.Add(pnlRxForm);

            Label lblRxFormTitle = new Label { Text = "🗑️ 처방내역 영구 삭제", Location = new Point(20, 15), Size = new Size(200, 25), Font = new Font("맑은 고딕", 11F, FontStyle.Bold), ForeColor = ColorAlarm };
            pnlRxForm.Controls.Add(lblRxFormTitle);

            Label lblRxFormWarning = new Label
            {
                Text = "※ 주의: 본 메뉴는 처방전 데이터(마스터 및 상세 내역)를 데이터베이스에서 물리적으로 완전히 삭제합니다.\n삭제 후에는 복구가 불가하므로 정확한 처방 건인지 확인 후 신중히 실행하십시오.",
                Location = new Point(20, 50),
                Size = new Size(280, 100),
                ForeColor = ColorAlarm,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            pnlRxForm.Controls.Add(lblRxFormWarning);

            _btnRxDeleteExecute = new Button
            {
                Text = "🗑️ 선택 처방 영구 삭제",
                Location = new Point(20, 170),
                Size = new Size(280, 40),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnRxDeleteExecute.FlatAppearance.BorderSize = 0;
            _btnRxDeleteExecute.Click += BtnRxDeleteExecute_Click;
            pnlRxForm.Controls.Add(_btnRxDeleteExecute);
        }

        // CRUD Event Handlers
        private void ClearUserForm()
        {
            _txtUserId.ReadOnly = false;
            _txtUserId.Text = "";
            _txtUserNm.Text = "";
            _txtUserPwd.Text = "";
            _txtUserDeptCd.Text = "";
            _txtUserLicNo.Text = "";
        }

        private void ClearCardForm()
        {
            _txtCardSlipSeq.Text = "";
            _txtCardRecpDt.Text = "";
            _txtCardChrtNo.Text = "";
            _txtCardCoNm.Text = "";
            _txtCardAmt.Text = "";
            _txtCardAdmNo.Text = "";
            _txtCardNo.Text = "";
        }

        private void ClearLabelForm()
        {
            _txtLabelDrugCode.ReadOnly = false;
            _txtLabelDrugCode.Text = "";
            _txtLabelDrug.Text = "";
            _txtLabelDan.Text = "";
            _txtLabelSave.Text = "";
            _txtLabelPrintOp.Text = "";
            _txtLabelInputOp.Text = "";
            _txtLabelEffct.Text = "";
            _txtLabelComment.Text = "";
            _txtLabelSampleUp.Text = "0";
            _txtLabelEffctUnit.Text = "";
        }

        private void DgvLabelInfos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = _dgvLabelInfos.Rows[e.RowIndex];
            _txtLabelDrugCode.Text = row.Cells["약품코드"].Value != null ? row.Cells["약품코드"].Value.ToString() : "";
            _txtLabelDrugCode.ReadOnly = true;
            _txtLabelDrug.Text = row.Cells["약품명"].Value != null ? row.Cells["약품명"].Value.ToString() : "";
            _txtLabelDan.Text = row.Cells["단위"].Value != null ? row.Cells["단위"].Value.ToString() : "";
            _txtLabelSave.Text = row.Cells["보관방법"].Value != null ? row.Cells["보관방법"].Value.ToString() : "";
            _txtLabelPrintOp.Text = row.Cells["출력옵션"].Value != null ? row.Cells["출력옵션"].Value.ToString() : "";
            _txtLabelInputOp.Text = row.Cells["입력옵션"].Value != null ? row.Cells["입력옵션"].Value.ToString() : "";
            _txtLabelEffct.Text = row.Cells["효능효과"].Value != null ? row.Cells["효능효과"].Value.ToString() : "";
            _txtLabelComment.Text = row.Cells["설명"].Value != null ? row.Cells["설명"].Value.ToString() : "";
            _txtLabelSampleUp.Text = row.Cells["샘플구분"].Value != null ? row.Cells["샘플구분"].Value.ToString() : "0";
            _txtLabelEffctUnit.Text = row.Cells["효능단위"].Value != null ? row.Cells["효능단위"].Value.ToString() : "";
        }

        private void DgvUsers_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = _dgvUsers.Rows[e.RowIndex];
            _txtUserId.Text = row.Cells["사용자 ID"].Value != null ? row.Cells["사용자 ID"].Value.ToString() : "";
            _txtUserId.ReadOnly = true;
            _txtUserNm.Text = row.Cells["이름"].Value != null ? row.Cells["이름"].Value.ToString() : "";
            _txtUserPwd.Text = "";
            _txtUserDeptCd.Text = row.Cells["부서 코드"].Value != null ? row.Cells["부서 코드"].Value.ToString() : "";
            _txtUserLicNo.Text = row.Cells["약사면허번호"].Value != null ? row.Cells["약사면허번호"].Value.ToString() : "";
        }

        private void DgvCardPays_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = _dgvCardPays.Rows[e.RowIndex];
            _txtCardSlipSeq.Text = row.Cells["일련번호"].Value != null ? row.Cells["일련번호"].Value.ToString() : "";
            _txtCardRecpDt.Text = row.Cells["수납일자"].Value != null ? row.Cells["수납일자"].Value.ToString() : "";
            _txtCardChrtNo.Text = row.Cells["차트번호"].Value != null ? row.Cells["차트번호"].Value.ToString() : "";
            _txtCardCoNm.Text = row.Cells["카드사명"].Value != null ? row.Cells["카드사명"].Value.ToString() : "";
            _txtCardAmt.Text = row.Cells["카드금액"].Value != null ? row.Cells["카드금액"].Value.ToString() : "";
            _txtCardAdmNo.Text = row.Cells["승인번호"].Value != null ? row.Cells["승인번호"].Value.ToString() : "";
            _txtCardNo.Text = row.Cells["카드번호"].Value != null ? row.Cells["카드번호"].Value.ToString() : "";
        }

        // ==========================================
        // SQL Service Control (MSSQL$PMPLUS20) Logic
        // ==========================================
        private string GetSqlServiceStatus()
        {
            if (_chkDemoMode.Checked)
            {
                if (_lastSqlServiceStatus == "UNKNOWN") _lastSqlServiceStatus = "RUNNING";
                return _lastSqlServiceStatus;
            }

            try
            {
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c sc query MSSQL$PMPLUS20",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = System.Text.Encoding.Default
                };
                using (System.Diagnostics.Process proc = System.Diagnostics.Process.Start(psi))
                {
                    if (proc == null) return "ERROR";
                    string output = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit();
                    if (output.Contains("RUNNING")) return "RUNNING";
                    if (output.Contains("STOPPED")) return "STOPPED";
                    if (output.Contains("1060")) return "NOT_INSTALLED";
                    if (output.Contains("STATE")) return "PENDING";
                    return "UNKNOWN";
                }
            }
            catch
            {
                return "ERROR";
            }
        }

        private void UpdateSqlServiceUI()
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            
            string status = GetSqlServiceStatus();
            this.BeginInvoke((Action)(() =>
            {
                switch (status)
                {
                    case "RUNNING":
                        _lblSqlServiceStatus.Text = "● SQL 서비스: 실행중";
                        _lblSqlServiceStatus.ForeColor = ColorEmerald;
                        _btnSqlServiceStart.Enabled = false;
                        _btnSqlServiceStop.Enabled = true;
                        break;
                    case "STOPPED":
                        _lblSqlServiceStatus.Text = "● SQL 서비스: 중지됨";
                        _lblSqlServiceStatus.ForeColor = ColorAlarm;
                        _btnSqlServiceStart.Enabled = true;
                        _btnSqlServiceStop.Enabled = false;
                        break;
                    case "PENDING":
                        _lblSqlServiceStatus.Text = "● SQL 서비스: 변경중...";
                        _lblSqlServiceStatus.ForeColor = Color.Orange;
                        _btnSqlServiceStart.Enabled = false;
                        _btnSqlServiceStop.Enabled = false;
                        break;
                    case "NOT_INSTALLED":
                        _lblSqlServiceStatus.Text = "● SQL 서비스: 미설치";
                        _lblSqlServiceStatus.ForeColor = Color.Gray;
                        _btnSqlServiceStart.Enabled = false;
                        _btnSqlServiceStop.Enabled = false;
                        break;
                    default:
                        _lblSqlServiceStatus.Text = "● SQL 서비스: 확인불가";
                        _lblSqlServiceStatus.ForeColor = Color.Gray;
                        _btnSqlServiceStart.Enabled = true;
                        _btnSqlServiceStop.Enabled = true;
                        break;
                }
            }));
        }

        private void ControlSqlService(bool start)
        {
            string cmd = start ? "start" : "stop";
            string title = start ? "서비스 시작" : "서비스 중지";
            string actionText = start ? "시작" : "중지";

            DialogResult dr = MessageBox.Show(
                string.Format("SQL Server (PMPLUS20) 서비스를 {0}하시겠습니까?\n실서버 모드 실행 시 관리자 권한 요청(UAC) 창이 뜰 수 있습니다.", actionText),
                title,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (dr != DialogResult.Yes) return;

            _lblSqlServiceStatus.Text = "● SQL 서비스: 변경중...";
            _lblSqlServiceStatus.ForeColor = Color.Orange;
            _btnSqlServiceStart.Enabled = false;
            _btnSqlServiceStop.Enabled = false;

            System.Threading.ThreadPool.QueueUserWorkItem(o =>
            {
                if (_chkDemoMode.Checked)
                {
                    System.Threading.Thread.Sleep(1500);
                    _lastSqlServiceStatus = start ? "RUNNING" : "STOPPED";
                    UpdateSqlServiceUI();
                    this.BeginInvoke((Action)(() =>
                    {
                        ShowToast(string.Format("SQL 서비스가 {0}되었습니다. (데모)", actionText), ColorEmerald);
                    }));
                }
                else
                {
                    try
                    {
                        System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "cmd.exe",
                            Arguments = string.Format("/c net {0} MSSQL$PMPLUS20", cmd),
                            UseShellExecute = true,
                            Verb = "runas",
                            WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                        };
                        using (System.Diagnostics.Process proc = System.Diagnostics.Process.Start(psi))
                        {
                            if (proc != null) proc.WaitForExit();
                        }
                    }
                    catch (Exception ex)
                    {
                        this.BeginInvoke((Action)(() =>
                        {
                            MessageBox.Show("서비스 제어 실패 (권한 거부 또는 취소됨):\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                    finally
                    {
                        System.Threading.Thread.Sleep(1500);
                        UpdateSqlServiceUI();
                    }
                }
            });
        }

        // ==========================================
        // Label Info (TBSIM040_43) CRUD Logic
        // ==========================================
        private void BtnLabelSearch_Click(object sender, EventArgs e)
        {
            string code = _txtLabelSearchCode.Text.Trim();
            string name = _txtLabelSearchName.Text.Trim();

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("약품코드");
                dt.Columns.Add("약품명");
                dt.Columns.Add("단위");
                dt.Columns.Add("보관방법");
                dt.Columns.Add("출력옵션");
                dt.Columns.Add("입력옵션");
                dt.Columns.Add("효능효과");
                dt.Columns.Add("설명");
                dt.Columns.Add("샘플구분");
                dt.Columns.Add("효능단위");

                foreach (var l in _mockLabelInfoList)
                {
                    bool match = true;
                    if (!string.IsNullOrEmpty(code) && !l.DrugCode.Contains(code)) match = false;
                    if (!string.IsNullOrEmpty(name) && !l.Drug.Contains(name)) match = false;
                    if (match)
                    {
                        dt.Rows.Add(l.DrugCode, l.Drug, l.Dan, l.Save, l.PrintOp, l.InputOp, l.Effct, l.Comment, l.SampleUp, l.EffctUnit);
                    }
                }
                _dgvLabelInfos.DataSource = dt;
                ShowToast(string.Format("라벨정보 {0}건 조회됨 (데모)", dt.Rows.Count), ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT LB_DRUGCODE AS [약품코드], LB_DRUG AS [약품명], LB_DAN AS [단위], LB_SAVE AS [보관방법], LB_PRINT_OP AS [출력옵션], LB_INPUT_OP AS [입력옵션], LB_EFFCT AS [효능효과], LB_COMMENT AS [설명], LB_SAMPLE_UP AS [샘플구분], LB_EFFCTUNIT AS [효능단위] FROM TBSIM040_43 WHERE 1=1";
                        if (!string.IsNullOrEmpty(code)) sql += " AND LB_DRUGCODE LIKE @code";
                        if (!string.IsNullOrEmpty(name)) sql += " AND LB_DRUG LIKE @name";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (!string.IsNullOrEmpty(code)) cmd.Parameters.AddWithValue("@code", "%" + code + "%");
                            if (!string.IsNullOrEmpty(name)) cmd.Parameters.AddWithValue("@name", "%" + name + "%");

                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    _dgvLabelInfos.DataSource = dt;
                    ShowToast(string.Format("라벨정보 {0}건 조회 완료", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("라벨정보 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnInventorySearch_Click(object sender, EventArgs e)
        {
            string keyword = _txtInventorySearch.Text.Trim();
            bool noNameOnly = _chkInventoryNoNameOnly.Checked;
            bool excludeZeroStock = _chkInventoryExcludeZeroStock != null && _chkInventoryExcludeZeroStock.Checked;

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("약품코드");
                dt.Columns.Add("약품명");
                dt.Columns.Add("제조회사");
                dt.Columns.Add("바코드");
                dt.Columns.Add("적정재고", typeof(decimal));
                dt.Columns.Add("재고합계", typeof(decimal));
                dt.Columns.Add("단가", typeof(decimal));
                dt.Columns.Add("재고금액(단가)", typeof(decimal));

                foreach (var item in _mockInventoryList)
                {
                    bool match = true;
                    if (!string.IsNullOrEmpty(keyword))
                    {
                        match = item.DrugCode.Contains(keyword) || (!string.IsNullOrEmpty(item.DrugName) && item.DrugName.Contains(keyword));
                    }
                    if (noNameOnly)
                    {
                        if (!string.IsNullOrEmpty(item.DrugName)) match = false;
                    }
                    if (excludeZeroStock && item.TotalStock == 0)
                    {
                        match = false;
                    }

                    if (match)
                    {
                        dt.Rows.Add(item.DrugCode, item.DrugName, item.Manufacturer, item.Barcode, item.ProperStock, item.TotalStock, item.UnitPrice, item.TotalStock * item.UnitPrice);
                    }
                }
                _dgvInventory.DataSource = dt;
                ShowToast(string.Format("재고 {0}건 조회됨 (데모)", dt.Rows.Count), ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT 
                                m.DRUG_CODE AS [약품코드],
                                m.ARTCNM AS [약품명],
                                m.MNF_CO_NM AS [제조회사],
                                COALESCE(s20.CD_CD_BARCODE, '') AS [바코드],
                                COALESCE(s20.CD_MY_UNIT, 0) AS [적정재고],
                                COALESCE(s8.MDCN_MQTY, 0) AS [재고합계],
                                COALESCE(s20.CD_IN_UNIT, 0) AS [단가],
                                CAST((COALESCE(s8.MDCN_MQTY, 0) * COALESCE(s20.CD_IN_UNIT, 0)) AS DECIMAL(18,0)) AS [재고금액(단가)]
                            FROM TBSIM040_01 m
                            LEFT JOIN TBSIM040_08 s8 ON m.DRUG_CODE = s8.DRUG_CODE
                            LEFT JOIN TBSIM040_20 s20 ON m.DRUG_CODE = s20.DRUG_CODE
                            WHERE 1=1";

                        if (!string.IsNullOrEmpty(keyword))
                        {
                            sql += " AND (m.DRUG_CODE LIKE @keyword OR m.ARTCNM LIKE @keyword)";
                        }
                        if (noNameOnly)
                        {
                            sql += " AND (m.ARTCNM IS NULL OR LTRIM(RTRIM(m.ARTCNM)) = '')";
                        }
                        if (excludeZeroStock)
                        {
                            sql += " AND COALESCE(s8.MDCN_MQTY, 0) <> 0";
                        }

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (!string.IsNullOrEmpty(keyword))
                            {
                                cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");
                            }

                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    _dgvInventory.DataSource = dt;
                    ShowToast(string.Format("재고 {0}건 조회됨", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("재고 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private DataTable CreateDurakanAuditTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("오류유형");
            dt.Columns.Add("테이블");
            dt.Columns.Add("식별번호");
            dt.Columns.Add("약품코드");
            dt.Columns.Add("일자");
            dt.Columns.Add("거래처/환자명");
            dt.Columns.Add("수량", typeof(decimal));
            dt.Columns.Add("기준");
            dt.Columns.Add("설명");
            return dt;
        }

        private string QuoteSqlName(string name)
        {
            return "[" + name.Replace("]", "]]") + "]";
        }

        private string FindFirstColumn(SqlConnection conn, string tableName, string[] candidates)
        {
            foreach (string candidate in candidates)
            {
                using (SqlCommand cmd = new SqlCommand(@"
                    SELECT TOP 1 c.name
                    FROM sys.columns c
                    INNER JOIN sys.objects o ON c.object_id = o.object_id
                    WHERE o.name = @tableName AND UPPER(c.name) = UPPER(@columnName);", conn))
                {
                    cmd.Parameters.AddWithValue("@tableName", tableName);
                    cmd.Parameters.AddWithValue("@columnName", candidate);
                    object found = cmd.ExecuteScalar();
                    if (found != null && found != DBNull.Value) return found.ToString();
                }
            }

            return "";
        }

        private List<string> GetTableColumnNames(SqlConnection conn, string tableName)
        {
            List<string> columns = new List<string>();
            using (SqlCommand cmd = new SqlCommand(@"
                SELECT c.name
                FROM sys.columns c
                INNER JOIN sys.objects o ON c.object_id = o.object_id
                WHERE o.name = @tableName
                ORDER BY c.column_id;", conn))
            {
                cmd.Parameters.AddWithValue("@tableName", tableName);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(reader.GetString(0));
                    }
                }
            }
            return columns;
        }

        private string FindColumnByPattern(SqlConnection conn, string tableName, string[] exactCandidates, string[] includeTokens, string[] excludeTokens)
        {
            string exact = FindFirstColumn(conn, tableName, exactCandidates);
            if (!string.IsNullOrEmpty(exact)) return exact;

            List<string> columns = GetTableColumnNames(conn, tableName);
            foreach (string col in columns)
            {
                string upper = col.ToUpper();
                bool include = includeTokens.Length == 0;
                foreach (string token in includeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        include = true;
                        break;
                    }
                }
                if (!include) continue;

                bool exclude = false;
                foreach (string token in excludeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        exclude = true;
                        break;
                    }
                }
                if (!exclude) return col;
            }

            return "";
        }

        private string FindColumnContainingValue(SqlConnection conn, string tableName, string value, string[] excludeTokens)
        {
            List<string> columns = GetTableColumnNames(conn, tableName);
            foreach (string col in columns)
            {
                string upper = col.ToUpper();
                bool exclude = false;
                foreach (string token in excludeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        exclude = true;
                        break;
                    }
                }
                if (exclude) continue;

                try
                {
                    string sql = "SELECT TOP 1 1 FROM " + QuoteSqlName(tableName) +
                                 " WHERE CONVERT(nvarchar(100), " + QuoteSqlName(col) + ") = @value";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@value", value);
                        object found = cmd.ExecuteScalar();
                        if (found != null && found != DBNull.Value) return col;
                    }
                }
                catch { }
            }

            return "";
        }

        private string FindSharedColumnByPattern(SqlConnection conn, string leftTable, string rightTable, string[] exactCandidates, string[] includeTokens, string[] excludeTokens)
        {
            List<string> leftColumns = GetTableColumnNames(conn, leftTable);
            List<string> rightColumns = GetTableColumnNames(conn, rightTable);

            foreach (string candidate in exactCandidates)
            {
                foreach (string leftCol in leftColumns)
                {
                    if (!leftCol.Equals(candidate, StringComparison.OrdinalIgnoreCase)) continue;
                    foreach (string rightCol in rightColumns)
                    {
                        if (rightCol.Equals(candidate, StringComparison.OrdinalIgnoreCase)) return leftCol;
                    }
                }
            }

            foreach (string leftCol in leftColumns)
            {
                string upper = leftCol.ToUpper();
                bool include = includeTokens.Length == 0;
                foreach (string token in includeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        include = true;
                        break;
                    }
                }
                if (!include) continue;

                bool exclude = false;
                foreach (string token in excludeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        exclude = true;
                        break;
                    }
                }
                if (exclude) continue;

                foreach (string rightCol in rightColumns)
                {
                    if (rightCol.Equals(leftCol, StringComparison.OrdinalIgnoreCase)) return leftCol;
                }
            }

            return "";
        }

        private string FindJoinedSampleTextColumn(SqlConnection conn, string detailTable, string masterTable, string joinCol, string detailCodeCol, string drugCode, string[] excludeTokens)
        {
            if (string.IsNullOrEmpty(joinCol) || string.IsNullOrEmpty(detailCodeCol)) return "";

            List<string> columns = new List<string>();
            using (SqlCommand cmd = new SqlCommand(@"
                SELECT c.name
                FROM sys.columns c
                INNER JOIN sys.objects o ON c.object_id = o.object_id
                INNER JOIN sys.types t ON c.user_type_id = t.user_type_id
                WHERE o.name = @tableName
                  AND t.name IN ('varchar', 'nvarchar', 'char', 'nchar')
                ORDER BY c.column_id;", conn))
            {
                cmd.Parameters.AddWithValue("@tableName", masterTable);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read()) columns.Add(reader.GetString(0));
                }
            }

            foreach (string col in columns)
            {
                string upper = col.ToUpper();
                bool exclude = false;
                foreach (string token in excludeTokens)
                {
                    if (upper.Contains(token.ToUpper()))
                    {
                        exclude = true;
                        break;
                    }
                }
                if (exclude) continue;

                try
                {
                    string sql = @"
                        SELECT TOP 1 CONVERT(nvarchar(200), m." + QuoteSqlName(col) + @")
                        FROM " + QuoteSqlName(detailTable) + @" d
                        INNER JOIN " + QuoteSqlName(masterTable) + @" m
                            ON CONVERT(nvarchar(100), d." + QuoteSqlName(joinCol) + @") = CONVERT(nvarchar(100), m." + QuoteSqlName(joinCol) + @")
                        WHERE CONVERT(nvarchar(30), d." + QuoteSqlName(detailCodeCol) + @") = @drugCode
                          AND NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(200), m." + QuoteSqlName(col) + @"))), '') IS NOT NULL
                          AND ISNUMERIC(CONVERT(nvarchar(200), m." + QuoteSqlName(col) + @")) = 0;";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@drugCode", drugCode);
                        object found = cmd.ExecuteScalar();
                        if (found != null && found != DBNull.Value) return col;
                    }
                }
                catch { }
            }

            return "";
        }

        private void BtnDurakanAudit_Click(object sender, EventArgs e)
        {
            RunStockMovementAudit(_dgvInventory, "644913503", 500m, 5m);
        }

        private void BtnStockAuditRun_Click(object sender, EventArgs e)
        {
            string drugCode = _txtStockAuditDrugCode.Text.Trim();
            decimal bottleUnit;
            decimal minPrescriptionQty;

            if (string.IsNullOrEmpty(drugCode))
            {
                MessageBox.Show("검사할 약품코드를 입력하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!decimal.TryParse(_txtStockAuditUnit.Text.Trim(), out bottleUnit) || bottleUnit <= 0)
            {
                MessageBox.Show("입고 기준단위는 0보다 큰 숫자로 입력하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!decimal.TryParse(_txtStockAuditMinQty.Text.Trim(), out minPrescriptionQty) || minPrescriptionQty < 0)
            {
                MessageBox.Show("최소 처방총량은 0 이상의 숫자로 입력하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            RunStockMovementAudit(_dgvStockMovementErrors, drugCode, bottleUnit, minPrescriptionQty);
        }

        private void BtnStockAuditDrugSearch_Click(object sender, EventArgs e)
        {
            string keyword = _txtStockAuditDrugName.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("검색할 약품명을 입력하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable dt = new DataTable();
            if (_chkDemoMode.Checked)
            {
                dt.Columns.Add("약품코드");
                dt.Columns.Add("약품명");
                dt.Columns.Add("청구단위");
                dt.Columns.Add("규격");
                dt.Columns.Add("적수");
                dt.Columns.Add("포장수량");
                dt.Rows.Add("644913503", "듀락칸이지시럽(락툴로오즈농축액)_(670g/500mL)", "500(1)mL", "500(1)", "1", "1");
                dt.Rows.Add("641601880", "알비스정_(1정)", "1정", "1정", "1", "1");
                _dgvStockAuditDrugSearch.DataSource = dt;
                ApplyContentSizedColumns(_dgvStockAuditDrugSearch);
                return;
            }

            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string billUnitCol = FindColumnByPattern(conn, "TBSIM040_01",
                        new string[] { "MDCN_UNIT", "BILL_UNIT", "CLAIM_UNIT", "CHARGE_UNIT", "CHRG_UNIT", "REQ_UNIT", "PAY_UNIT", "DRUG_UNIT", "DRUG_DAN", "DAN", "UNIT", "청구단위" },
                        new string[] { "UNIT", "DAN", "청구", "단위" },
                        new string[] { "PRICE", "COST", "MONEY", "AMT", "QTY", "CNT", "CODE", "CD", "DATE", "DTIME" });
                    string standardCol = FindColumnByPattern(conn, "TBSIM040_01",
                        new string[] { "MDCN_STANDARD", "MDCN_STAND", "DRUG_SPEC", "DRUG_STANDARD", "STANDARD", "STAND", "SPEC", "SIZE", "규격" },
                        new string[] { "STANDARD", "STAND", "SPEC", "SIZE", "규격" },
                        new string[] { "PRICE", "COST", "MONEY", "AMT", "QTY", "CNT", "CODE", "CD", "DATE", "DTIME" });
                    string packCol = FindColumnByPattern(conn, "TBSIM040_20",
                        new string[] { "JEOKSU", "JUCKSU", "JUKSU", "CD_MY_UNIT", "CD_JUKSU", "CD_JEOKSU", "CD_JUCKSU", "PACK_CNT", "PACK_QTY", "UNIT_CNT", "UNIT_QTY", "적수" },
                        new string[] { "JEOK", "JUCK", "JUK", "UNIT", "적수" },
                        new string[] { "PRICE", "COST", "MONEY", "AMT", "DANGA", "BAR", "CODE", "CD", "DATE", "DTIME", "IN_UNIT" });
                    string packQtyCol = FindColumnByPattern(conn, "TBSIM040_20",
                        new string[] { "CD_PACK_QTY", "PACK_QTY", "PACKAGE_QTY", "BOX_QTY", "포장수량" },
                        new string[] { "PACK", "PACKAGE", "BOX", "포장" },
                        new string[] { "PRICE", "COST", "MONEY", "AMT", "DANGA", "BAR", "CODE", "CD", "DATE", "DTIME" });
                    string nameSpecExpr = @"CASE 
                                WHEN CHARINDEX('_(', m.ARTCNM) > 0 AND CHARINDEX(')', m.ARTCNM, CHARINDEX('_(', m.ARTCNM) + 2) > 0
                                THEN SUBSTRING(m.ARTCNM, CHARINDEX('_(', m.ARTCNM) + 2, CHARINDEX(')', m.ARTCNM, CHARINDEX('_(', m.ARTCNM) + 2) - CHARINDEX('_(', m.ARTCNM) - 2)
                                ELSE ''
                            END";
                    string billUnitExpr = string.IsNullOrEmpty(billUnitCol) ? nameSpecExpr : "COALESCE(NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), m." + QuoteSqlName(billUnitCol) + "))), ''), " + nameSpecExpr + ")";
                    string standardExpr = string.IsNullOrEmpty(standardCol) ? nameSpecExpr : "COALESCE(NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), m." + QuoteSqlName(standardCol) + "))), ''), " + nameSpecExpr + ")";
                    string packExpr = string.IsNullOrEmpty(packCol) ? "'1'" : "COALESCE(NULLIF(LTRIM(RTRIM(CONVERT(nvarchar(100), s20." + QuoteSqlName(packCol) + "))), ''), '1')";
                    string packQtyExpr = string.IsNullOrEmpty(packQtyCol) ? "''" : "COALESCE(CONVERT(nvarchar(100), s20." + QuoteSqlName(packQtyCol) + "), '')";

                    string sql = @"
                        SELECT TOP 50
                            m.DRUG_CODE AS [약품코드],
                            m.ARTCNM AS [약품명],
                            COALESCE(m.MNF_CO_NM, '') AS [제조회사],
                            " + billUnitExpr + @" AS [청구단위],
                            " + standardExpr + @" AS [규격],
                            " + packExpr + @" AS [적수],
                            " + packQtyExpr + @" AS [포장수량],
                            COALESCE(s8.MDCN_MQTY, 0) AS [재고량]
                        FROM TBSIM040_01 m
                        LEFT JOIN TBSIM040_20 s20 ON m.DRUG_CODE = s20.DRUG_CODE
                        LEFT JOIN TBSIM040_08 s8 ON m.DRUG_CODE = s8.DRUG_CODE
                        WHERE m.ARTCNM LIKE @keyword OR m.DRUG_CODE LIKE @keyword
                        ORDER BY m.ARTCNM, m.DRUG_CODE;";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }

                _dgvStockAuditDrugSearch.DataSource = dt;
                ApplyContentSizedColumns(_dgvStockAuditDrugSearch);
                ShowToast(string.Format("약품 {0}건 검색됨", dt.Rows.Count), ColorEmerald);
            }
            catch (Exception ex)
            {
                MessageBox.Show("약품 검색 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void DgvStockAuditDrugSearch_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || _dgvStockAuditDrugSearch == null) return;
            DataGridViewRow row = _dgvStockAuditDrugSearch.Rows[e.RowIndex];

            string code = GetGridCellText(row, "약품코드");
            if (!string.IsNullOrEmpty(code)) _txtStockAuditDrugCode.Text = code;

            string name = GetGridCellText(row, "약품명");
            string billUnit = GetGridCellText(row, "청구단위");
            string standard = GetGridCellText(row, "규격");
            string pack = GetGridCellText(row, "적수");
            string packQty = GetGridCellText(row, "포장수량");
            string stock = GetGridCellText(row, "재고량");
            _txtStockAuditDrugInfo.Text =
                "약품명: " + name + Environment.NewLine +
                "약품코드: " + code + Environment.NewLine +
                "청구단위: " + billUnit + "    규격: " + standard + Environment.NewLine +
                "적수: " + pack + "    포장수량: " + packQty + "    재고량: " + stock;
        }

        private void RunStockMovementAudit(DataGridView targetGrid, string drugCode, decimal bottleUnit, decimal minPrescriptionQty)
        {
            if (targetGrid == null) return;

            DataTable result = CreateDurakanAuditTable();

            if (_chkDemoMode.Checked)
            {
                result.Rows.Add("입고 500배수 아님", "TBSWH040_02", "DEMO-IN-001", drugCode, "2026-06-19", "태전약품판매주식회사", 150m, "500mL 병 제품은 500의 배수 입고", "입고수량 150은 500mL 병 단위와 맞지 않습니다.");
                result.Rows.Add("처방 총투여량 5 미만", "TBSID040_04", "DEMO-RX-001", drugCode, "", "", 3m, "총투여량 5 이상", "총투여량이 5mL 미만인 처방입니다.");
                targetGrid.DataSource = result;
                ApplyContentSizedColumns(targetGrid);
                ShowToast("입출고 오류 검사 완료 (데모)", ColorEmerald);
                return;
            }

            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    string inboundTable = "TBSWH040_02";
                    string inboundMasterTable = "TBSWH040_01";
                    string codeCol = FindColumnByPattern(conn, inboundTable,
                        new string[] { "DRUG_CODE", "MDCN_CD", "ITEM_CD", "GOODS_CODE", "CD_CODE", "CD", "DRUGCD", "MDCNCD" },
                        new string[] { "DRUG", "MDCN", "GOOD", "ITEM", "CD", "CODE" },
                        new string[] { "BAR", "NAME", "NM", "DATE", "DTE", "DTIME" });
                    if (string.IsNullOrEmpty(codeCol))
                    {
                        codeCol = FindColumnContainingValue(conn, inboundTable, drugCode, new string[] { "BAR", "NAME", "NM", "DATE", "DTE", "DTIME", "QTY", "AMT", "PRICE", "COST", "MONEY" });
                    }

                    string qtyCol = FindColumnByPattern(conn, inboundTable,
                        new string[] { "QTY", "IN_QTY", "INPUT_QTY", "PUR_QTY", "DRUG_QTY", "MDCN_QTY", "SU", "IN_SU", "CNT", "QNT", "IN_QNT", "MDCN_QNT" },
                        new string[] { "QTY", "QNT", "SU", "CNT" },
                        new string[] { "PRICE", "COST", "MONEY", "AMT", "UNIT", "DANGA", "TAX", "VAT", "TOTAL", "SUM", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });
                    string idCol = FindColumnByPattern(conn, inboundTable,
                        new string[] { "IN_SEQ", "SEQ", "INPUT_SEQ", "SLIP_SEQ", "ROW_ID", "IDX" },
                        new string[] { "SEQ", "IDX", "NO", "NUM" },
                        new string[] { "CODE", "CD", "DATE", "DTE", "DTIME" });

                    if (!string.IsNullOrEmpty(codeCol) && !string.IsNullOrEmpty(qtyCol))
                    {
                        string joinCol = FindSharedColumnByPattern(conn, inboundTable, inboundMasterTable,
                            new string[] { "IN_SEQ", "SEQ", "INPUT_SEQ", "SLIP_SEQ", "PRES_SEQ", "ROW_ID", "IDX", "IN_NO", "INPUT_NO", "SLIP_NO", "PRES_DTIME" },
                            new string[] { "PRES", "SEQ", "IDX", "NO", "NUM", "DTIME", "DATE", "DT" },
                            new string[] { "DRUG", "MDCN", "GOOD", "ITEM", "CODE", "CD", "QTY", "AMT", "PRICE", "COST", "MONEY" });
                        string idExpr = string.IsNullOrEmpty(idCol) ? "''" : "CONVERT(nvarchar(50), d." + QuoteSqlName(idCol) + ")";
                        string packCol = FindColumnByPattern(conn, inboundTable,
                            new string[] { "PACK_QTY", "PACK_CNT", "PACK_SU", "JEOKSU", "JUCKSU", "JUKSU", "JQTY", "BOX_QTY", "BOX_CNT", "UNIT_QTY", "UNIT_CNT" },
                            new string[] { "PACK", "JEOK", "JUCK", "JUK", "BOX", "UNIT" },
                            new string[] { "PRICE", "COST", "MONEY", "AMT", "DANGA", "TAX", "VAT", "TOTAL", "SUM", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });
                        string packExpr = string.IsNullOrEmpty(packCol) ? "1" : "CAST(d." + QuoteSqlName(packCol) + " AS decimal(18,4))";
                        string qtyExpr = "CAST(d." + QuoteSqlName(qtyCol) + " AS decimal(18,4))";
                        string inboundQtyExpr = "(" + packExpr + " * " + qtyExpr + ")";
                        string dateCol = FindColumnByPattern(conn, inboundTable,
                            new string[] { "PRES_DTIME", "IN_DATE", "IN_DT", "INPUT_DATE", "INPUT_DT", "PUR_DATE", "PUR_DT", "SLIP_DATE", "SLIP_DT" },
                            new string[] { "DATE", "DTIME", "DT", "DTE", "DAY" },
                            new string[] { "CODE", "CD", "QTY", "AMT", "PRICE", "COST" });
                        string vendorCol = FindColumnByPattern(conn, inboundTable,
                            new string[] { "CUST_NM", "CUST_NAME", "CUSTNM", "CUSTOMER_NM", "CUSTOMER_NAME", "CLIENT_NM", "CLIENT_NAME", "VENDOR_NM", "VENDOR_NAME", "VEND_NM", "VEND_NAME", "SUPPLIER_NM", "SUPPLIER_NAME", "SUPPLY_NM", "TRD_NM", "TRADER_NM", "TRADE_NM", "TR_NM", "BP_NM", "PARTNER_NM", "ACCT_NM" },
                            new string[] { "CUST", "CLIENT", "VENDOR", "VEND", "SUPPL", "SUPPLY", "TRD", "TRADE", "COMP", "CORP", "BUSI", "BP", "PARTNER", "ACCT" },
                            new string[] { "CODE", "CD", "NO", "SEQ", "DATE", "DTIME", "QTY", "AMT", "PRICE", "COST", "MNF", "MANUF", "MAKER" });
                        string masterDateCol = FindColumnByPattern(conn, inboundMasterTable,
                            new string[] { "PRES_DTIME", "IN_DATE", "IN_DT", "INPUT_DATE", "INPUT_DT", "PUR_DATE", "PUR_DT", "SLIP_DATE", "SLIP_DT" },
                            new string[] { "DATE", "DTIME", "DT", "DTE", "DAY" },
                            new string[] { "CODE", "CD", "QTY", "AMT", "PRICE", "COST" });
                        string masterVendorCol = FindColumnByPattern(conn, inboundMasterTable,
                            new string[] { "CUST_NM", "CUST_NAME", "CUSTNM", "CUSTOMER_NM", "CUSTOMER_NAME", "CLIENT_NM", "CLIENT_NAME", "VENDOR_NM", "VENDOR_NAME", "VEND_NM", "VEND_NAME", "SUPPLIER_NM", "SUPPLIER_NAME", "SUPPLY_NM", "TRD_NM", "TRADER_NM", "TRADE_NM", "TR_NM", "BP_NM", "PARTNER_NM", "ACCT_NM", "BUSI_NM", "COMP_NM", "CORP_NM", "CMP_NM", "CUST", "CUSTOMER", "VENDOR", "VEND", "SUPPLIER", "TRADER", "거래처", "매입처", "구입처", "공급처", "업체명" },
                            new string[] { "CUST", "CLIENT", "VENDOR", "VEND", "SUPPL", "SUPPLY", "TRD", "TRADE", "COMP", "CORP", "BUSI", "CMP", "BP", "PARTNER", "ACCT", "거래", "매입", "구입", "공급", "업체" },
                            new string[] { "CODE", "CD", "NO", "SEQ", "DATE", "DTIME", "QTY", "AMT", "PRICE", "COST", "MNF", "MANUF", "MAKER" });
                        string fromClause = "FROM " + QuoteSqlName(inboundTable) + " d";
                        bool hasMasterJoin = !string.IsNullOrEmpty(joinCol);
                        if (hasMasterJoin)
                        {
                            fromClause += " LEFT JOIN " + QuoteSqlName(inboundMasterTable) + " m ON CONVERT(nvarchar(100), d." + QuoteSqlName(joinCol) + ") = CONVERT(nvarchar(100), m." + QuoteSqlName(joinCol) + ")";
                        }
                        if (string.IsNullOrEmpty(masterVendorCol) && hasMasterJoin)
                        {
                            masterVendorCol = FindJoinedSampleTextColumn(conn, inboundTable, inboundMasterTable, joinCol, codeCol, drugCode,
                                new string[] { "CODE", "CD", "NO", "SEQ", "DATE", "DTIME", "QTY", "AMT", "PRICE", "COST", "MNF", "MANUF", "MAKER", "MEMO", "REMARK", "NOTE", "REPORT", "STATE", "STATUS", "TYPE", "KIND", "GUBUN", "GBN", "구분", "메모", "비고", "보고", "상태" });
                        }
                        string dateExpr = !string.IsNullOrEmpty(dateCol)
                            ? "CONVERT(nvarchar(30), d." + QuoteSqlName(dateCol) + ")"
                            : (!string.IsNullOrEmpty(masterDateCol) && hasMasterJoin ? "CONVERT(nvarchar(30), m." + QuoteSqlName(masterDateCol) + ")" : "''");
                        string vendorExpr = !string.IsNullOrEmpty(vendorCol)
                            ? "CONVERT(nvarchar(100), d." + QuoteSqlName(vendorCol) + ")"
                            : (!string.IsNullOrEmpty(masterVendorCol) && hasMasterJoin ? "CONVERT(nvarchar(100), m." + QuoteSqlName(masterVendorCol) + ")" : "''");
                        string sqlInbound = @"
                            SELECT TOP 1000
                                '입고 500배수 아님' AS [오류유형],
                                'TBSWH040_02' AS [테이블],
                                " + idExpr + @" AS [식별번호],
                                CONVERT(nvarchar(30), d." + QuoteSqlName(codeCol) + @") AS [약품코드],
                                " + dateExpr + @" AS [일자],
                                " + vendorExpr + @" AS [거래처/환자명],
                                " + inboundQtyExpr + @" AS [수량],
                                '500mL 병 제품은 500의 배수 입고' AS [기준],
                                '적수 × 수량 기준 입고량이 500의 배수가 아닙니다. 병 단위 입고 후 1mL 단위 수불 관리 여부를 확인하세요.' AS [설명]
                            " + fromClause + @"
                            WHERE CONVERT(nvarchar(30), d." + QuoteSqlName(codeCol) + @") = @drugCode
                              AND " + inboundQtyExpr + @" <> 0
                              AND (" + inboundQtyExpr + @" - FLOOR(" + inboundQtyExpr + @" / @bottleUnit) * @bottleUnit) <> 0;";

                        using (SqlCommand cmd = new SqlCommand(sqlInbound, conn))
                        {
                            cmd.Parameters.AddWithValue("@drugCode", drugCode);
                            cmd.Parameters.AddWithValue("@bottleUnit", bottleUnit);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(result);
                            }
                        }
                    }
                    else
                    {
                        string cols = string.Join(", ", GetTableColumnNames(conn, inboundTable).ToArray());
                        result.Rows.Add("입고 검사 불가", inboundTable, "", drugCode, "", "", 0m, "약품코드/수량 컬럼 자동 탐색", "입고장 상세 테이블의 약품코드 또는 수량 컬럼명을 찾지 못했습니다. 실제 컬럼: " + cols);
                    }

                    string rxTable = "TBSID040_04";
                    string rxMasterTable = "TBSID040_03";
                    string rxCodeCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "DRUG_CODE", "MDCN_CD", "ITEM_CD", "GOODS_CODE", "CD_CODE", "CD", "DRUGCD", "MDCNCD" },
                        new string[] { "DRUG", "MDCN", "GOOD", "ITEM", "CD", "CODE" },
                        new string[] { "BAR", "NAME", "NM", "DATE", "DTE", "DTIME" });
                    if (string.IsNullOrEmpty(rxCodeCol))
                    {
                        rxCodeCol = FindColumnContainingValue(conn, rxTable, drugCode, new string[] { "BAR", "NAME", "NM", "DATE", "DTE", "DTIME", "QTY", "AMT", "PRICE", "COST", "MONEY" });
                    }
                    string rxQtyCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "DRUG_QTY", "QTY", "MDCN_QTY", "TOT_QTY", "TOTAL_QTY", "DOSAGE_QTY", "USE_QTY", "DAY_QTY", "INPUT_QTY", "TQTY", "TOTQTY", "CNT", "SU", "QNT", "TOT_QNT", "TOTAL_QNT" },
                        new string[] { "QTY", "QNT", "SU", "CNT", "AMT", "DOS", "DAY", "TOT", "TOTAL" },
                        new string[] { "PRICE", "COST", "MONEY", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });
                    string rxSeqCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "DRUG_SEQ", "PRES_SEQ", "RX_SEQ", "SEQ", "SLIP_SEQ", "IDX" },
                        new string[] { "SEQ", "IDX", "NO", "NUM" },
                        new string[] { "CODE", "CD", "DATE", "DTE", "DTIME" });
                    string rxDoseCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "ONCE_QTY", "ONCE_QNT", "ONE_QTY", "ONE_QNT", "ONE_TIME_QTY", "PER_QTY", "DOSE_QTY", "DOSAGE_QTY", "DAY_QTY", "QTY_ONCE", "QTY1", "QTY_1", "DOSAGE", "1회량" },
                        new string[] { "ONCE", "ONE", "DOSE", "DOSAGE", "PER", "QTY1", "1회" },
                        new string[] { "TOT", "TOTAL", "SUM", "CNT", "DAY", "DAYS", "PRICE", "COST", "MONEY", "AMT", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });
                    string rxTimesCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "TIMES", "TIME_CNT", "CNT", "DAY_CNT", "USE_CNT", "DOSAGE_CNT", "FREQ", "FREQ_CNT", "BOK_CNT", "HOESU", "HWE_CNT", "횟수" },
                        new string[] { "CNT", "FREQ", "TIME", "HOE", "HWE", "횟" },
                        new string[] { "TOT", "TOTAL", "SUM", "DAY_QTY", "QTY", "PRICE", "COST", "MONEY", "AMT", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });
                    string rxDaysCol = FindColumnByPattern(conn, rxTable,
                        new string[] { "DAYS", "DAY_CNT", "PRES_DAYS", "DRUG_DAYS", "DUR_DAYS", "ILSU", "일수" },
                        new string[] { "DAY", "DAYS", "ILSU", "일수" },
                        new string[] { "QTY", "AMT", "PRICE", "COST", "MONEY", "BAR", "CODE", "CD", "DATE", "DTE", "DTIME" });

                    if (!string.IsNullOrEmpty(rxCodeCol) && (!string.IsNullOrEmpty(rxQtyCol) || (!string.IsNullOrEmpty(rxDoseCol) && !string.IsNullOrEmpty(rxTimesCol) && !string.IsNullOrEmpty(rxDaysCol))))
                    {
                        string rxJoinCol = FindSharedColumnByPattern(conn, rxTable, rxMasterTable,
                            new string[] { "DRUG_SEQ", "PRES_SEQ", "RX_SEQ", "SLIP_SEQ", "SEQ", "IDX" },
                            new string[] { "SEQ", "IDX", "NO", "NUM" },
                            new string[] { "CODE", "CD", "DATE", "DTE", "DTIME", "QTY", "AMT", "PRICE", "COST" });
                        string rxDateCol = FindColumnByPattern(conn, rxTable,
                            new string[] { "PRES_DTIME", "PRES_DT", "MED_YMD", "MED_DT", "SUNAB_DT", "RCPT_DT", "RECP_DT", "DATE", "DT" },
                            new string[] { "DATE", "DTIME", "DT", "DTE", "DAY", "YMD" },
                            new string[] { "CODE", "CD", "QTY", "AMT", "PRICE", "COST" });
                        string rxMasterDateCol = FindColumnByPattern(conn, rxMasterTable,
                            new string[] { "PRES_DTIME", "PRES_DT", "MED_YMD", "MED_DT", "SUNAB_DT", "RCPT_DT", "RECP_DT", "DATE", "DT" },
                            new string[] { "DATE", "DTIME", "DT", "DTE", "DAY", "YMD" },
                            new string[] { "CODE", "CD", "QTY", "AMT", "PRICE", "COST" });
                        string rxNameCol = FindColumnByPattern(conn, rxTable,
                            new string[] { "PAT_NM", "PAT_NAME", "PT_NM", "PT_NAME", "PATIENT_NM", "PATIENT_NAME", "CUST_NM", "NAME", "NM", "환자명", "이름" },
                            new string[] { "PAT", "PT", "PATIENT", "CUST", "NAME", "NM", "환자", "이름" },
                            new string[] { "DRUG", "MDCN", "GOOD", "ITEM", "CODE", "CD", "DATE", "DTIME", "QTY", "AMT", "PRICE", "COST" });
                        string rxMasterNameCol = FindColumnByPattern(conn, rxMasterTable,
                            new string[] { "PAT_NM", "PAT_NAME", "PT_NM", "PT_NAME", "PATIENT_NM", "PATIENT_NAME", "CUST_NM", "NAME", "NM", "환자명", "이름" },
                            new string[] { "PAT", "PT", "PATIENT", "CUST", "NAME", "NM", "환자", "이름" },
                            new string[] { "DRUG", "MDCN", "GOOD", "ITEM", "CODE", "CD", "DATE", "DTIME", "QTY", "AMT", "PRICE", "COST" });
                        string rxFromClause = "FROM " + QuoteSqlName(rxTable) + " r";
                        bool hasRxMasterJoin = !string.IsNullOrEmpty(rxJoinCol);
                        if (hasRxMasterJoin)
                        {
                            rxFromClause += " LEFT JOIN " + QuoteSqlName(rxMasterTable) + " p ON CONVERT(nvarchar(100), r." + QuoteSqlName(rxJoinCol) + ") = CONVERT(nvarchar(100), p." + QuoteSqlName(rxJoinCol) + ")";
                        }
                        string rxIdExpr = string.IsNullOrEmpty(rxSeqCol) ? "''" : "CONVERT(nvarchar(50), r." + QuoteSqlName(rxSeqCol) + ")";
                        string rxQtyExpr = (!string.IsNullOrEmpty(rxDoseCol) && !string.IsNullOrEmpty(rxTimesCol) && !string.IsNullOrEmpty(rxDaysCol))
                            ? "(CAST(r." + QuoteSqlName(rxDoseCol) + " AS decimal(18,4)) * CAST(r." + QuoteSqlName(rxTimesCol) + " AS decimal(18,4)) * CAST(r." + QuoteSqlName(rxDaysCol) + " AS decimal(18,4)))"
                            : "CAST(r." + QuoteSqlName(rxQtyCol) + " AS decimal(18,4))";
                        string rxDateExpr = !string.IsNullOrEmpty(rxDateCol)
                            ? "CONVERT(nvarchar(30), r." + QuoteSqlName(rxDateCol) + ")"
                            : (!string.IsNullOrEmpty(rxMasterDateCol) && hasRxMasterJoin ? "CONVERT(nvarchar(30), p." + QuoteSqlName(rxMasterDateCol) + ")" : "''");
                        string rxNameExpr = !string.IsNullOrEmpty(rxNameCol)
                            ? "CONVERT(nvarchar(100), r." + QuoteSqlName(rxNameCol) + ")"
                            : (!string.IsNullOrEmpty(rxMasterNameCol) && hasRxMasterJoin ? "CONVERT(nvarchar(100), p." + QuoteSqlName(rxMasterNameCol) + ")" : "''");
                        string sqlRx = @"
                            SELECT TOP 1000
                                '처방 총투여량 5 미만' AS [오류유형],
                                'TBSID040_04' AS [테이블],
                                " + rxIdExpr + @" AS [식별번호],
                                CONVERT(nvarchar(30), r." + QuoteSqlName(rxCodeCol) + @") AS [약품코드],
                                " + rxDateExpr + @" AS [일자],
                                " + rxNameExpr + @" AS [거래처/환자명],
                                " + rxQtyExpr + @" AS [수량],
                                '총투여량 5 이상' AS [기준],
                                '듀락칸 500mL 소분 제품의 총투여량이 5mL 미만입니다. 처방 입력 단위/수량을 확인하세요.' AS [설명]
                            " + rxFromClause + @"
                            WHERE CONVERT(nvarchar(30), r." + QuoteSqlName(rxCodeCol) + @") = @drugCode
                              AND " + rxQtyExpr + @" < @minQty;";

                        using (SqlCommand cmd = new SqlCommand(sqlRx, conn))
                        {
                            cmd.Parameters.AddWithValue("@drugCode", drugCode);
                            cmd.Parameters.AddWithValue("@minQty", minPrescriptionQty);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(result);
                            }
                        }
                    }
                    else
                    {
                        string cols = string.Join(", ", GetTableColumnNames(conn, "TBSID040_04").ToArray());
                        result.Rows.Add("처방 검사 불가", "TBSID040_04", "", drugCode, "", "", 0m, "약품코드/총투여량 컬럼 자동 탐색", "처방 상세 테이블의 약품코드 또는 총투여량 컬럼명을 찾지 못했습니다. 실제 컬럼: " + cols);
                    }
                }

                targetGrid.DataSource = result;
                ApplyContentSizedColumns(targetGrid);
                ShowToast(string.Format("입출고 오류 후보 {0}건 조회됨", result.Rows.Count), ColorEmerald);
            }
            catch (Exception ex)
            {
                MessageBox.Show("입출고 오류 검사 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private string GetGridCellText(DataGridViewRow row, params string[] columnNames)
        {
            if (row == null || row.DataGridView == null) return "";

            foreach (string name in columnNames)
            {
                if (row.DataGridView.Columns.Contains(name))
                {
                    object value = row.Cells[name].Value;
                    return value != null ? value.ToString() : "";
                }
            }

            foreach (DataGridViewColumn col in row.DataGridView.Columns)
            {
                foreach (string name in columnNames)
                {
                    if (string.Equals(col.HeaderText, name, StringComparison.OrdinalIgnoreCase))
                    {
                        object value = row.Cells[col.Index].Value;
                        return value != null ? value.ToString() : "";
                    }
                }
            }

            return "";
        }

        private void DgvInventory_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridViewRow row = _dgvInventory.Rows[e.RowIndex];

            _txtInvFormDrugCode.Text = GetGridCellText(row, "약품코드");
            _txtInvFormBarcode.Text = GetGridCellText(row, "바코드");
            _txtInvFormDrugName.Text = GetGridCellText(row, "약품명");
            _txtInvFormManufacturer.Text = GetGridCellText(row, "제조회사");

            string barcode = _txtInvFormBarcode.Text;
            _lblInvFormSuggest.Text = GetBarcodeSuggestion(barcode);
        }

        private string GetBarcodeSuggestion(string barcode)
        {
            if (string.IsNullOrEmpty(barcode)) return "";
            if (barcode == "8801328109268") return "💡 추천명: 에프킬라 모기향 캔 (40개입)";
            if (barcode == "8806113706554") return "💡 추천명: 에프킬라 살충제 (SC존슨)";
            if (barcode == "8809004779903") return "💡 추천명: 홈키파/홈매트 살충제 (헨켈)";
            if (barcode == "8806573019812") return "💡 추천명: 대웅바이오 의약품";
            
            if (barcode.StartsWith("8806113")) return "💡 추천 제조사: 에스씨존슨코리아 (에프킬라)";
            if (barcode.StartsWith("8809004")) return "💡 추천 제조사: 헨켈홈케어코리아 (홈매트/홈키파)";
            if (barcode.StartsWith("8806573")) return "💡 추천 제조사: 대웅바이오 (의약품)";
            
            return "";
        }

        private void BtnInvFormUpdate_Click(object sender, EventArgs e)
        {
            string code = _txtInvFormDrugCode.Text.Trim();
            string name = _txtInvFormDrugName.Text.Trim();
            string manufacturer = _txtInvFormManufacturer.Text.Trim();

            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("먼저 수정할 약품을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("약품명은 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                var item = _mockInventoryList.Find(x => x.DrugCode == code);
                if (item != null)
                {
                    item.DrugName = name;
                    item.Manufacturer = manufacturer;
                    ShowToast("약품 정보가 수정되었습니다. (데모)", ColorEmerald);
                    BtnInventorySearch_Click(null, null);
                }
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "UPDATE TBSIM040_01 SET ARTCNM = @name, MNF_CO_NM = @manufacturer WHERE DRUG_CODE = @code";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@code", code);
                            cmd.Parameters.AddWithValue("@name", name);
                            cmd.Parameters.AddWithValue("@manufacturer", string.IsNullOrEmpty(manufacturer) ? (object)DBNull.Value : manufacturer);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("약품 정보가 성공적으로 수정되었습니다.", ColorEmerald);
                    BtnInventorySearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("약품 정보 수정 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnInvFormDelete_Click(object sender, EventArgs e)
        {
            string code = _txtInvFormDrugCode.Text.Trim();
            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("먼저 삭제할 약품을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal stock = 0;
            if (_dgvInventory.CurrentRow != null)
            {
                stock = Convert.ToDecimal(_dgvInventory.CurrentRow.Cells["재고합계"].Value);
            }

            if (stock != 0)
            {
                MessageBox.Show("재고가 0이 아닌 약품은 임의로 삭제할 수 없습니다.", "삭제 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(string.Format("약품코드 [{0}] 약품을 정말로 데이터베이스에서 삭제하시겠습니까?", code), "약품 개별 삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                _mockInventoryList.RemoveAll(x => x.DrugCode == code);
                ShowToast("약품이 삭제되었습니다. (데모)", ColorEmerald);
                ClearInvForm();
                BtnInventorySearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                string sql8 = "DELETE FROM TBSIM040_08 WHERE DRUG_CODE = @code";
                                string sql20 = "DELETE FROM TBSIM040_20 WHERE DRUG_CODE = @code";
                                string sql01 = "DELETE FROM TBSIM040_01 WHERE DRUG_CODE = @code";

                                using (SqlCommand cmd = new SqlCommand(sql8, conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@code", code);
                                    cmd.ExecuteNonQuery();
                                }
                                using (SqlCommand cmd = new SqlCommand(sql20, conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@code", code);
                                    cmd.ExecuteNonQuery();
                                }
                                using (SqlCommand cmd = new SqlCommand(sql01, conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@code", code);
                                    cmd.ExecuteNonQuery();
                                }
                                trans.Commit();
                            }
                            catch (Exception)
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }
                    ShowToast("약품이 성공적으로 삭제되었습니다.", ColorEmerald);
                    ClearInvForm();
                    BtnInventorySearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("약품 삭제 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ClearInvForm()
        {
            _txtInvFormDrugCode.Text = "";
            _txtInvFormBarcode.Text = "";
            _txtInvFormDrugName.Text = "";
            _txtInvFormManufacturer.Text = "";
            _lblInvFormSuggest.Text = "";
        }

        private void BtnInvBatchDelete_Click(object sender, EventArgs e)
        {
            if (_chkDemoMode.Checked)
            {
                int deletedCount = _mockInventoryList.RemoveAll(x => string.IsNullOrEmpty(x.DrugName) && x.TotalStock == 0);
                MessageBox.Show(string.Format("[데모] 이름이 없고 재고가 0인 약품 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearInvForm();
                BtnInventorySearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                int candidateCount = 0;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string countSql = @"
                            SELECT COUNT(*) 
                            FROM TBSIM040_01 m
                            LEFT JOIN TBSIM040_08 s8 ON m.DRUG_CODE = s8.DRUG_CODE
                            WHERE (m.ARTCNM IS NULL OR LTRIM(RTRIM(m.ARTCNM)) = '')
                              AND (COALESCE(s8.MDCN_MQTY, 0) = 0);";
                        
                        using (SqlCommand cmd = new SqlCommand(countSql, conn))
                        {
                            candidateCount = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (candidateCount == 0)
                {
                    MessageBox.Show("삭제 대상인 '이름이 없고 재고가 0인 약품'이 존재하지 않습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult dr = MessageBox.Show(
                    string.Format("이름이 등록되지 않았고 재고가 0인 약품 {0}건을 정말로 영구 일괄 삭제하시겠습니까?\n\n" +
                                  "※ 주의: 삭제 완료 후 복구할 수 없습니다.", candidateCount),
                    "이름 없는 재고 0 약품 일괄 삭제 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (dr != DialogResult.Yes) return;

                this.Cursor = Cursors.WaitCursor;
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        int deletedCount = 0;
                        using (SqlConnection conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            using (SqlTransaction trans = conn.BeginTransaction())
                            {
                                try
                                {
                                    List<string> codesToDelete = new List<string>();
                                    string selectSql = @"
                                        SELECT m.DRUG_CODE 
                                        FROM TBSIM040_01 m
                                        LEFT JOIN TBSIM040_08 s8 ON m.DRUG_CODE = s8.DRUG_CODE
                                        WHERE (m.ARTCNM IS NULL OR LTRIM(RTRIM(m.ARTCNM)) = '')
                                          AND (COALESCE(s8.MDCN_MQTY, 0) = 0);";
                                    
                                    using (SqlCommand selectCmd = new SqlCommand(selectSql, conn, trans))
                                    {
                                        using (SqlDataReader reader = selectCmd.ExecuteReader())
                                        {
                                            while (reader.Read())
                                            {
                                                codesToDelete.Add(reader.GetString(0));
                                            }
                                        }
                                    }

                                    if (codesToDelete.Count > 0)
                                    {
                                        foreach (string code in codesToDelete)
                                        {
                                            using (SqlCommand cmd = new SqlCommand("DELETE FROM TBSIM040_08 WHERE DRUG_CODE = @code", conn, trans))
                                            {
                                                cmd.Parameters.AddWithValue("@code", code);
                                                cmd.ExecuteNonQuery();
                                            }
                                            using (SqlCommand cmd = new SqlCommand("DELETE FROM TBSIM040_20 WHERE DRUG_CODE = @code", conn, trans))
                                            {
                                                cmd.Parameters.AddWithValue("@code", code);
                                                cmd.ExecuteNonQuery();
                                            }
                                            using (SqlCommand cmd = new SqlCommand("DELETE FROM TBSIM040_01 WHERE DRUG_CODE = @code", conn, trans))
                                            {
                                                cmd.Parameters.AddWithValue("@code", code);
                                                cmd.ExecuteNonQuery();
                                            }
                                            deletedCount++;
                                        }
                                    }
                                    trans.Commit();
                                }
                                catch (Exception)
                                {
                                    trans.Rollback();
                                    throw;
                                }
                            }
                        }

                        this.BeginInvoke((Action)(() =>
                        {
                            this.Cursor = Cursors.Default;
                            MessageBox.Show(string.Format("이름 없는 재고 0 약품 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ClearInvForm();
                            BtnInventorySearch_Click(null, null);
                        }));
                    }
                    catch (Exception ex)
                    {
                        this.BeginInvoke((Action)(() =>
                        {
                            this.Cursor = Cursors.Default;
                            MessageBox.Show("약품 삭제 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                });
            }
        }

        private void BtnInvCleanDupBarcodes_Click(object sender, EventArgs e)
        {
            if (_chkDemoMode.Checked)
            {
                int deletedCount = 0;
                var nameless = _mockInventoryList.Where(x => string.IsNullOrEmpty(x.DrugName)).ToList();
                var groups = nameless.GroupBy(x => new { x.DrugCode, x.Barcode });
                
                foreach (var g in groups)
                {
                    if (g.Count() > 1)
                    {
                        var list = g.ToList();
                        for (int i = 0; i < list.Count - 1; i++)
                        {
                            _mockInventoryList.Remove(list[i]);
                            deletedCount++;
                        }
                    }
                }

                MessageBox.Show(string.Format("[데모] 중복된 바코드 매핑 {0}건이 정상적으로 정리되었습니다.", deletedCount), "정리 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearInvForm();
                BtnInventorySearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                int duplicateCount = 0;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string countSql = @"
                            WITH DuplicateBarcodes AS (
                                SELECT DRUG_CODE, CD_CD_BARCODE, SN,
                                       ROW_NUMBER() OVER (
                                           PARTITION BY DRUG_CODE, CD_CD_BARCODE 
                                           ORDER BY SN DESC
                                       ) as rn
                                FROM TBSIM040_20
                                WHERE DRUG_CODE IN (SELECT DRUG_CODE FROM TBSIM040_01 WHERE ARTCNM IS NULL OR LTRIM(RTRIM(ARTCNM)) = '')
                            )
                            SELECT COUNT(*) FROM DuplicateBarcodes WHERE rn > 1;";
                        
                        using (SqlCommand cmd = new SqlCommand(countSql, conn))
                        {
                            duplicateCount = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("중복 바코드 조회 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (duplicateCount == 0)
                {
                    MessageBox.Show("정리 대상인 '이름 없는 약품의 중복 바코드 매핑'이 존재하지 않습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult dr = MessageBox.Show(
                    string.Format("이름이 없는 약품 중 중복 등록된 바코드 매핑 {0}건을 정리하시겠습니까?\n\n" +
                                  "※ 가장 최신에 등록/수정된 바코드 매핑 1건만 유지되고 나머지는 영구 삭제됩니다.", duplicateCount),
                    "중복 바코드 매핑 정리 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );

                if (dr != DialogResult.Yes) return;

                this.Cursor = Cursors.WaitCursor;
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        int deletedCount = 0;
                        using (SqlConnection conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            string cleanDeleteSql = @"
                                WITH DuplicateBarcodes AS (
                                    SELECT DRUG_CODE, CD_CD_BARCODE, SN,
                                           ROW_NUMBER() OVER (
                                               PARTITION BY DRUG_CODE, CD_CD_BARCODE 
                                               ORDER BY SN DESC
                                           ) as rn
                                    FROM TBSIM040_20
                                    WHERE DRUG_CODE IN (SELECT DRUG_CODE FROM TBSIM040_01 WHERE ARTCNM IS NULL OR LTRIM(RTRIM(ARTCNM)) = '')
                                )
                                DELETE FROM TBSIM040_20
                                WHERE SN IN (
                                    SELECT SN 
                                    FROM DuplicateBarcodes 
                                    WHERE rn > 1
                                );";

                            using (SqlCommand cmd = new SqlCommand(cleanDeleteSql, conn))
                            {
                                cmd.CommandTimeout = 300;
                                deletedCount = cmd.ExecuteNonQuery();
                            }
                        }

                        this.BeginInvoke((Action)(() =>
                        {
                            this.Cursor = Cursors.Default;
                            MessageBox.Show(string.Format("중복 바코드 매핑 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "정리 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ClearInvForm();
                            BtnInventorySearch_Click(null, null);
                        }));
                    }
                    catch (Exception ex)
                    {
                        this.BeginInvoke((Action)(() =>
                        {
                            this.Cursor = Cursors.Default;
                            MessageBox.Show("중복 바코드 정리 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                });
            }
        }

        private void BtnInvBarcodeSearchWeb_Click(object sender, EventArgs e)
        {
            string barcode = _txtInvFormBarcode.Text.Trim();
            if (string.IsNullOrEmpty(barcode))
            {
                MessageBox.Show("조회할 바코드가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start("https://gs1.koreannet.or.kr/pr/" + barcode);
            }
            catch (Exception ex)
            {
                MessageBox.Show("웹 브라우저를 열지 못했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLabelAdd_Click(object sender, EventArgs e)
        {
            string code = _txtLabelDrugCode.Text.Trim();
            string drug = _txtLabelDrug.Text.Trim();
            string dan = _txtLabelDan.Text.Trim();
            string save = _txtLabelSave.Text.Trim();
            string print = _txtLabelPrintOp.Text.Trim();
            string input = _txtLabelInputOp.Text.Trim();
            string effct = _txtLabelEffct.Text.Trim();
            string comment = _txtLabelComment.Text.Trim();
            string sample = _txtLabelSampleUp.Text.Trim();
            string unit = _txtLabelEffctUnit.Text.Trim();

            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("약품코드는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(sample)) sample = "0";

            if (_chkDemoMode.Checked)
            {
                if (_mockLabelInfoList.Exists(l => l.DrugCode == code))
                {
                    MessageBox.Show("이미 존재하는 약품코드입니다.", "중복 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                _mockLabelInfoList.Add(new MockLabelInfo { DrugCode = code, Drug = drug, Dan = dan, Save = save, PrintOp = print, InputOp = input, Effct = effct, Comment = comment, SampleUp = sample, EffctUnit = unit });
                ShowToast("라벨정보가 추가되었습니다. (데모)", ColorEmerald);
                ClearLabelForm();
                BtnLabelSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string checkSql = "SELECT COUNT(*) FROM TBSIM040_43 WHERE LB_DRUGCODE = @code";
                        using (SqlCommand checkCmd = new SqlCommand(checkSql, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@code", code);
                            if (Convert.ToInt32(checkCmd.ExecuteScalar()) > 0)
                            {
                                MessageBox.Show("이미 존재하는 약품코드입니다.", "중복 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        string sql = "INSERT INTO TBSIM040_43 (LB_DRUGCODE, LB_DRUG, LB_DAN, LB_SAVE, LB_PRINT_OP, LB_INPUT_OP, LB_EFFCT, LB_COMMENT, LB_SAMPLE_UP, LB_EFFCTUNIT) VALUES (@code, @drug, @dan, @save, @print, @input, @effct, @comment, @sample, @unit)";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@code", code);
                            cmd.Parameters.AddWithValue("@drug", string.IsNullOrEmpty(drug) ? (object)DBNull.Value : drug);
                            cmd.Parameters.AddWithValue("@dan", string.IsNullOrEmpty(dan) ? (object)DBNull.Value : dan);
                            cmd.Parameters.AddWithValue("@save", string.IsNullOrEmpty(save) ? (object)DBNull.Value : save);
                            cmd.Parameters.AddWithValue("@print", string.IsNullOrEmpty(print) ? (object)DBNull.Value : print);
                            cmd.Parameters.AddWithValue("@input", string.IsNullOrEmpty(input) ? (object)DBNull.Value : input);
                            cmd.Parameters.AddWithValue("@effct", string.IsNullOrEmpty(effct) ? (object)DBNull.Value : effct);
                            cmd.Parameters.AddWithValue("@comment", string.IsNullOrEmpty(comment) ? (object)DBNull.Value : comment);
                            cmd.Parameters.AddWithValue("@sample", sample);
                            cmd.Parameters.AddWithValue("@unit", string.IsNullOrEmpty(unit) ? (object)DBNull.Value : unit);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("라벨정보가 성공적으로 추가되었습니다.", ColorEmerald);
                    ClearLabelForm();
                    BtnLabelSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("라벨정보 추가 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnLabelUpdate_Click(object sender, EventArgs e)
        {
            string code = _txtLabelDrugCode.Text.Trim();
            string drug = _txtLabelDrug.Text.Trim();
            string dan = _txtLabelDan.Text.Trim();
            string save = _txtLabelSave.Text.Trim();
            string print = _txtLabelPrintOp.Text.Trim();
            string input = _txtLabelInputOp.Text.Trim();
            string effct = _txtLabelEffct.Text.Trim();
            string comment = _txtLabelComment.Text.Trim();
            string sample = _txtLabelSampleUp.Text.Trim();
            string unit = _txtLabelEffctUnit.Text.Trim();

            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("수정할 대상을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrEmpty(sample)) sample = "0";

            if (_chkDemoMode.Checked)
            {
                var l = _mockLabelInfoList.Find(x => x.DrugCode == code);
                if (l == null)
                {
                    MessageBox.Show("수정할 대상을 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                l.Drug = drug;
                l.Dan = dan;
                l.Save = save;
                l.PrintOp = print;
                l.InputOp = input;
                l.Effct = effct;
                l.Comment = comment;
                l.SampleUp = sample;
                l.EffctUnit = unit;

                ShowToast("라벨정보가 수정되었습니다. (데모)", ColorEmerald);
                ClearLabelForm();
                BtnLabelSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "UPDATE TBSIM040_43 SET LB_DRUG = @drug, LB_DAN = @dan, LB_SAVE = @save, LB_PRINT_OP = @print, LB_INPUT_OP = @input, LB_EFFCT = @effct, LB_COMMENT = @comment, LB_SAMPLE_UP = @sample, LB_EFFCTUNIT = @unit WHERE LB_DRUGCODE = @code";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@code", code);
                            cmd.Parameters.AddWithValue("@drug", string.IsNullOrEmpty(drug) ? (object)DBNull.Value : drug);
                            cmd.Parameters.AddWithValue("@dan", string.IsNullOrEmpty(dan) ? (object)DBNull.Value : dan);
                            cmd.Parameters.AddWithValue("@save", string.IsNullOrEmpty(save) ? (object)DBNull.Value : save);
                            cmd.Parameters.AddWithValue("@print", string.IsNullOrEmpty(print) ? (object)DBNull.Value : print);
                            cmd.Parameters.AddWithValue("@input", string.IsNullOrEmpty(input) ? (object)DBNull.Value : input);
                            cmd.Parameters.AddWithValue("@effct", string.IsNullOrEmpty(effct) ? (object)DBNull.Value : effct);
                            cmd.Parameters.AddWithValue("@comment", string.IsNullOrEmpty(comment) ? (object)DBNull.Value : comment);
                            cmd.Parameters.AddWithValue("@sample", sample);
                            cmd.Parameters.AddWithValue("@unit", string.IsNullOrEmpty(unit) ? (object)DBNull.Value : unit);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("라벨정보가 성공적으로 수정되었습니다.", ColorEmerald);
                    ClearLabelForm();
                    BtnLabelSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("라벨정보 수정 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnLabelDelete_Click(object sender, EventArgs e)
        {
            string code = _txtLabelDrugCode.Text.Trim();
            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("삭제할 대상을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(string.Format("약품코드 [{0}] 라벨 정보를 영구 삭제하시겠습니까?", code), "삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                _mockLabelInfoList.RemoveAll(x => x.DrugCode == code);
                ShowToast("라벨정보가 삭제되었습니다. (데모)", ColorEmerald);
                ClearLabelForm();
                BtnLabelSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "DELETE FROM TBSIM040_43 WHERE LB_DRUGCODE = @code";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@code", code);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("라벨정보가 성공적으로 삭제되었습니다.", ColorEmerald);
                    ClearLabelForm();
                    BtnLabelSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("라벨정보 삭제 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // ==========================================
        // Prescription Delete (TBSID040_03, 04, 05) Logic
        // ==========================================
        private void BtnRxDelSearch_Click(object sender, EventArgs e)
        {
            string name = _txtRxDelSearchName.Text.Trim();
            string jumin = _txtRxDelSearchJumin.Text.Trim();

            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(jumin))
            {
                MessageBox.Show("환자명 또는 주민번호 중 하나는 입력해야 검색이 가능합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("처방일련번호");
                dt.Columns.Add("처방일시");
                dt.Columns.Add("환자명");
                dt.Columns.Add("주민번호");
                dt.Columns.Add("수납일자");

                foreach (var rx in _mockPrescriptionList)
                {
                    bool match = true;
                    if (!string.IsNullOrEmpty(name) && !rx.PatNm.Contains(name)) match = false;
                    if (!string.IsNullOrEmpty(jumin) && !rx.PatJuminNo.Contains(jumin)) match = false;
                    if (match)
                    {
                        dt.Rows.Add(rx.DrugSeq, rx.PresDtime, rx.PatNm, rx.PatJuminNo, rx.SunabDt);
                    }
                }
                _dgvRxDeleteList.DataSource = dt;
                ShowToast(string.Format("처방내역 {0}건 조회됨 (데모)", dt.Rows.Count), ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT DRUG_SEQ AS [처방일련번호], PRES_DTIME AS [처방일시], PAT_NM AS [환자명], PAT_JUMIN_NO AS [주민번호], SUNAB_DT AS [수납일자] FROM TBSID040_03 WHERE 1=1";
                        if (!string.IsNullOrEmpty(name)) sql += " AND PAT_NM LIKE @name";
                        if (!string.IsNullOrEmpty(jumin)) sql += " AND PAT_JUMIN_NO LIKE @jumin";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (!string.IsNullOrEmpty(name)) cmd.Parameters.AddWithValue("@name", "%" + name + "%");
                            if (!string.IsNullOrEmpty(jumin)) cmd.Parameters.AddWithValue("@jumin", "%" + jumin + "%");

                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    _dgvRxDeleteList.DataSource = dt;
                    ShowToast(string.Format("처방내역 {0}건 조회 완료", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("처방내역 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnRxDeleteExecute_Click(object sender, EventArgs e)
        {
            if (_dgvRxDeleteList.SelectedRows.Count == 0)
            {
                MessageBox.Show("삭제할 처방 행을 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var row = _dgvRxDeleteList.SelectedRows[0];
            string seq = row.Cells["처방일련번호"].Value != null ? row.Cells["처방일련번호"].Value.ToString() : "";
            string patName = row.Cells["환자명"].Value != null ? row.Cells["환자명"].Value.ToString() : "";

            if (string.IsNullOrEmpty(seq))
            {
                MessageBox.Show("선택된 행의 처방일련번호(DRUG_SEQ)가 유효하지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 1차 확인
            DialogResult dr1 = MessageBox.Show(
                string.Format("정말로 환자 [{0}]의 처방전 데이터(처방일련번호: {1})를 영구 삭제하시겠습니까?", patName, seq),
                "처방전 영구 삭제 - 1차 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );
            if (dr1 != DialogResult.Yes) return;

            // 2차 확인
            DialogResult dr2 = MessageBox.Show(
                "이 작업은 실제 데이터베이스 테이블(TBSID040_04, TBSID040_05, TBSID040_03)의 모든 연관 레코드를 물리적으로 제거합니다.\n삭제 후에는 절대 복구할 수 없습니다. 계속하시겠습니까?",
                "처방전 영구 삭제 - 최종 경고 (복구 불가)",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Stop
            );
            if (dr2 != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                _mockPrescriptionList.RemoveAll(x => x.DrugSeq == seq);
                ShowToast("처방전 내역이 삭제되었습니다. (데모)", ColorEmerald);
                BtnRxDelSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                // 1. TBSID040_04 상세 삭제
                                string sql4 = "DELETE FROM TBSID040_04 WHERE DRUG_SEQ = @seq";
                                using (SqlCommand cmd4 = new SqlCommand(sql4, conn, trans))
                                {
                                    cmd4.Parameters.AddWithValue("@seq", seq);
                                    cmd4.ExecuteNonQuery();
                                }

                                // 2. TBSID040_05 상세 삭제
                                string sql5 = "DELETE FROM TBSID040_05 WHERE DRUG_SEQ = @seq";
                                using (SqlCommand cmd5 = new SqlCommand(sql5, conn, trans))
                                {
                                    cmd5.Parameters.AddWithValue("@seq", seq);
                                    cmd5.ExecuteNonQuery();
                                }

                                // 3. TBSID040_03 마스터 삭제
                                string sql3 = "DELETE FROM TBSID040_03 WHERE DRUG_SEQ = @seq";
                                using (SqlCommand cmd3 = new SqlCommand(sql3, conn, trans))
                                {
                                    cmd3.Parameters.AddWithValue("@seq", seq);
                                    cmd3.ExecuteNonQuery();
                                }

                                trans.Commit();
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }
                    ShowToast("처방 내역이 데이터베이스에서 성공적으로 영구 삭제되었습니다.", ColorEmerald);
                    BtnRxDelSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("처방전 내역 삭제 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnUserSearch_Click(object sender, EventArgs e)
        {
            string id = _txtUserSearchId.Text.Trim();
            string name = _txtUserSearchName.Text.Trim();

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("사용자 ID");
                dt.Columns.Add("이름");
                dt.Columns.Add("부서 코드");
                dt.Columns.Add("약사면허번호");

                foreach (var u in _mockUserList)
                {
                    bool match = true;
                    if (!string.IsNullOrEmpty(id) && !u.UserId.Contains(id)) match = false;
                    if (!string.IsNullOrEmpty(name) && !u.UserNm.Contains(name)) match = false;
                    if (match)
                    {
                        dt.Rows.Add(u.UserId, u.UserNm, u.DeptCd, u.LicNo);
                    }
                }
                _dgvUsers.DataSource = dt;
                ShowToast(string.Format("사용자 {0}명 조회됨 (데모)", dt.Rows.Count), ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT user_id AS [사용자 ID], user_nm AS [이름], dept_cd AS [부서 코드], lic_no AS [약사면허번호] FROM TBSIM000_09 WHERE 1=1";
                        if (!string.IsNullOrEmpty(id)) sql += " AND user_id LIKE @user_id";
                        if (!string.IsNullOrEmpty(name)) sql += " AND user_nm LIKE @user_nm";
                        
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (!string.IsNullOrEmpty(id)) cmd.Parameters.AddWithValue("@user_id", "%" + id + "%");
                            if (!string.IsNullOrEmpty(name)) cmd.Parameters.AddWithValue("@user_nm", "%" + name + "%");
                            
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    _dgvUsers.DataSource = dt;
                    ShowToast(string.Format("사용자 {0}명 조회됨", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("사용자 ID 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnUserAdd_Click(object sender, EventArgs e)
        {
            string id = _txtUserId.Text.Trim();
            string nm = _txtUserNm.Text.Trim();
            string pwd = _txtUserPwd.Text;
            string dept = _txtUserDeptCd.Text.Trim();
            string lic = _txtUserLicNo.Text.Trim();

            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nm) || string.IsNullOrEmpty(pwd))
            {
                MessageBox.Show("사용자 ID, 이름, 비밀번호는 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string pwdHash = GetSHA512Hash(pwd);

            if (_chkDemoMode.Checked)
            {
                if (_mockUserList.Exists(u => u.UserId == id))
                {
                    MessageBox.Show("이미 존재하는 사용자 ID입니다.", "중복 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                _mockUserList.Add(new MockUser { UserId = id, UserNm = nm, UserPwd = pwdHash, DeptCd = dept, LicNo = lic });
                ShowToast("사용자가 추가되었습니다. (데모)", ColorEmerald);
                ClearUserForm();
                BtnUserSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        // 중복체크
                        string checkSql = "SELECT COUNT(*) FROM TBSIM000_09 WHERE user_id = @id";
                        using (SqlCommand checkCmd = new SqlCommand(checkSql, conn))
                        {
                            checkCmd.Parameters.AddWithValue("@id", id);
                            if (Convert.ToInt32(checkCmd.ExecuteScalar()) > 0)
                            {
                                MessageBox.Show("이미 존재하는 사용자 ID입니다.", "중복 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        string sql = "INSERT INTO TBSIM000_09 (user_id, user_nm, user_pwd, dept_cd, lic_no) VALUES (@id, @nm, @pwd, @dept, @lic)";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@id", id);
                            cmd.Parameters.AddWithValue("@nm", nm);
                            cmd.Parameters.AddWithValue("@pwd", pwdHash);
                            cmd.Parameters.AddWithValue("@dept", dept);
                            cmd.Parameters.AddWithValue("@lic", lic);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("사용자가 성공적으로 추가되었습니다.", ColorEmerald);
                    ClearUserForm();
                    BtnUserSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("사용자 추가 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnUserUpdate_Click(object sender, EventArgs e)
        {
            string id = _txtUserId.Text.Trim();
            string nm = _txtUserNm.Text.Trim();
            string pwd = _txtUserPwd.Text;
            string dept = _txtUserDeptCd.Text.Trim();
            string lic = _txtUserLicNo.Text.Trim();

            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nm))
            {
                MessageBox.Show("사용자 ID와 이름은 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                var u = _mockUserList.Find(user => user.UserId == id);
                if (u == null)
                {
                    MessageBox.Show("수정할 사용자를 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                u.UserNm = nm;
                if (!string.IsNullOrEmpty(pwd))
                {
                    u.UserPwd = GetSHA512Hash(pwd);
                }
                u.DeptCd = dept;
                u.LicNo = lic;

                ShowToast("사용자 정보가 수정되었습니다. (데모)", ColorEmerald);
                ClearUserForm();
                BtnUserSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "";
                        if (!string.IsNullOrEmpty(pwd))
                        {
                            sql = "UPDATE TBSIM000_09 SET user_nm = @nm, user_pwd = @pwd, dept_cd = @dept, lic_no = @lic WHERE user_id = @id";
                        }
                        else
                        {
                            sql = "UPDATE TBSIM000_09 SET user_nm = @nm, dept_cd = @dept, lic_no = @lic WHERE user_id = @id";
                        }

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@id", id);
                            cmd.Parameters.AddWithValue("@nm", nm);
                            if (!string.IsNullOrEmpty(pwd))
                            {
                                cmd.Parameters.AddWithValue("@pwd", GetSHA512Hash(pwd));
                            }
                            cmd.Parameters.AddWithValue("@dept", dept);
                            cmd.Parameters.AddWithValue("@lic", lic);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("사용자 정보가 성공적으로 수정되었습니다.", ColorEmerald);
                    ClearUserForm();
                    BtnUserSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("사용자 수정 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnUserDelete_Click(object sender, EventArgs e)
        {
            string id = _txtUserId.Text.Trim();
            if (string.IsNullOrEmpty(id))
            {
                MessageBox.Show("삭제할 사용자를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(string.Format("사용자 '{0}'를 정말 삭제하시겠습니까?", id), "삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                _mockUserList.RemoveAll(u => u.UserId == id);
                ShowToast("사용자가 삭제되었습니다. (데모)", ColorEmerald);
                ClearUserForm();
                BtnUserSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "DELETE FROM TBSIM000_09 WHERE user_id = @id";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@id", id);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("사용자가 성공적으로 삭제되었습니다.", ColorEmerald);
                    ClearUserForm();
                    BtnUserSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("사용자 삭제 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnCardSearch_Click(object sender, EventArgs e)
        {
            string chart = _txtCardSearchChart.Text.Trim();
            string date = _txtCardSearchDate.Text.Trim();

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("일련번호", typeof(decimal));
                dt.Columns.Add("수납일자");
                dt.Columns.Add("차트번호");
                dt.Columns.Add("카드사명");
                dt.Columns.Add("카드금액", typeof(decimal));
                dt.Columns.Add("승인번호");
                dt.Columns.Add("카드번호");

                foreach (var c in _mockCardPayList)
                {
                    bool match = true;
                    if (!string.IsNullOrEmpty(chart) && !c.ChrtNo.Contains(chart)) match = false;
                    if (!string.IsNullOrEmpty(date) && !c.RecpDt.Contains(date)) match = false;
                    if (match)
                    {
                        dt.Rows.Add(c.SlipSeq, c.RecpDt, c.ChrtNo, c.CardCoNm, c.CardAmt, c.CardAdmNo, c.CardNo);
                    }
                }
                _dgvCardPays.DataSource = dt;
                ShowToast(string.Format("카드결제 {0}건 조회됨 (데모)", dt.Rows.Count), ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT slip_seq AS [일련번호], recp_dt AS [수납일자], chrtno AS [차트번호], card_co_nm AS [카드사명], card_amt AS [카드금액], card_adm_no AS [승인번호], card_no AS [카드번호] FROM tbsir000_01 WHERE 1=1";
                        if (!string.IsNullOrEmpty(chart)) sql += " AND chrtno LIKE @chart";
                        if (!string.IsNullOrEmpty(date)) sql += " AND recp_dt LIKE @date";
                        
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (!string.IsNullOrEmpty(chart)) cmd.Parameters.AddWithValue("@chart", "%" + chart + "%");
                            if (!string.IsNullOrEmpty(date)) cmd.Parameters.AddWithValue("@date", "%" + date + "%");
                            
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                    _dgvCardPays.DataSource = dt;
                    ShowToast(string.Format("카드결제 {0}건 조회됨", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("카드결제내역 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnCardAdd_Click(object sender, EventArgs e)
        {
            string recpDt = _txtCardRecpDt.Text.Trim();
            string chrtNo = _txtCardChrtNo.Text.Trim();
            string cardCo = _txtCardCoNm.Text.Trim();
            string amtStr = _txtCardAmt.Text.Trim();
            string admNo = _txtCardAdmNo.Text.Trim();
            string cardNo = _txtCardNo.Text.Trim();

            if (string.IsNullOrEmpty(recpDt) || string.IsNullOrEmpty(chrtNo) || string.IsNullOrEmpty(amtStr))
            {
                MessageBox.Show("수납일자, 차트번호, 카드금액은 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal amt;
            if (!decimal.TryParse(amtStr, out amt))
            {
                MessageBox.Show("올바른 카드금액(숫자)을 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                decimal nextSeq = 10001;
                if (_mockCardPayList.Count > 0)
                {
                    nextSeq = _mockCardPayList.Max(c => c.SlipSeq) + 1;
                }
                _mockCardPayList.Add(new MockCardPay
                {
                    SlipSeq = nextSeq,
                    RecpDt = recpDt,
                    ChrtNo = chrtNo,
                    CardCoNm = cardCo,
                    CardAmt = amt,
                    CardAdmNo = admNo,
                    CardNo = cardNo
                });
                ShowToast("카드결제내역이 추가되었습니다. (데모)", ColorEmerald);
                ClearCardForm();
                BtnCardSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        decimal nextSeq = 1;
                        string getSeqSql = "SELECT ISNULL(MAX(slip_seq), 0) + 1 FROM tbsir000_01";
                        using (SqlCommand seqCmd = new SqlCommand(getSeqSql, conn))
                        {
                            nextSeq = Convert.ToDecimal(seqCmd.ExecuteScalar());
                        }

                        string sql = "INSERT INTO tbsir000_01 (slip_seq, recp_dt, chrtno, card_co_nm, card_amt, card_adm_no, card_no) VALUES (@seq, @recp_dt, @chrtno, @card_co_nm, @card_amt, @card_adm_no, @card_no)";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@seq", nextSeq);
                            cmd.Parameters.AddWithValue("@recp_dt", recpDt);
                            cmd.Parameters.AddWithValue("@chrtno", chrtNo);
                            cmd.Parameters.AddWithValue("@card_co_nm", cardCo);
                            cmd.Parameters.AddWithValue("@card_amt", amt);
                            cmd.Parameters.AddWithValue("@card_adm_no", admNo);
                            cmd.Parameters.AddWithValue("@card_no", cardNo);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("카드결제내역이 성공적으로 추가되었습니다.", ColorEmerald);
                    ClearCardForm();
                    BtnCardSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("카드결제내역 추가 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnCardUpdate_Click(object sender, EventArgs e)
        {
            string seqStr = _txtCardSlipSeq.Text.Trim();
            string recpDt = _txtCardRecpDt.Text.Trim();
            string chrtNo = _txtCardChrtNo.Text.Trim();
            string cardCo = _txtCardCoNm.Text.Trim();
            string amtStr = _txtCardAmt.Text.Trim();
            string admNo = _txtCardAdmNo.Text.Trim();
            string cardNo = _txtCardNo.Text.Trim();

            if (string.IsNullOrEmpty(seqStr))
            {
                MessageBox.Show("수정할 결제내역을 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(recpDt) || string.IsNullOrEmpty(chrtNo) || string.IsNullOrEmpty(amtStr))
            {
                MessageBox.Show("수납일자, 차트번호, 카드금액은 필수 입력 항목입니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal seq = decimal.Parse(seqStr);
            decimal amt;
            if (!decimal.TryParse(amtStr, out amt))
            {
                MessageBox.Show("올바른 카드금액(숫자)을 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_chkDemoMode.Checked)
            {
                var c = _mockCardPayList.Find(card => card.SlipSeq == seq);
                if (c == null)
                {
                    MessageBox.Show("수정할 결제내역을 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                c.RecpDt = recpDt;
                c.ChrtNo = chrtNo;
                c.CardCoNm = cardCo;
                c.CardAmt = amt;
                c.CardAdmNo = admNo;
                c.CardNo = cardNo;

                ShowToast("카드결제내역이 수정되었습니다. (데모)", ColorEmerald);
                ClearCardForm();
                BtnCardSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "UPDATE tbsir000_01 SET recp_dt = @recp_dt, chrtno = @chrtno, card_co_nm = @card_co_nm, card_amt = @card_amt, card_adm_no = @card_adm_no, card_no = @card_no WHERE slip_seq = @seq";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@seq", seq);
                            cmd.Parameters.AddWithValue("@recp_dt", recpDt);
                            cmd.Parameters.AddWithValue("@chrtno", chrtNo);
                            cmd.Parameters.AddWithValue("@card_co_nm", cardCo);
                            cmd.Parameters.AddWithValue("@card_amt", amt);
                            cmd.Parameters.AddWithValue("@card_adm_no", admNo);
                            cmd.Parameters.AddWithValue("@card_no", cardNo);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("카드결제내역이 성공적으로 수정되었습니다.", ColorEmerald);
                    ClearCardForm();
                    BtnCardSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("카드결제내역 수정 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnCardDelete_Click(object sender, EventArgs e)
        {
            string seqStr = _txtCardSlipSeq.Text.Trim();
            if (string.IsNullOrEmpty(seqStr))
            {
                MessageBox.Show("삭제할 결제내역을 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            decimal seq = decimal.Parse(seqStr);

            DialogResult dr = MessageBox.Show(string.Format("일련번호 '{0}'인 결제내역을 정말 삭제하시겠습니까?", seq), "삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                _mockCardPayList.RemoveAll(c => c.SlipSeq == seq);
                ShowToast("카드결제내역이 삭제되었습니다. (데모)", ColorEmerald);
                ClearCardForm();
                BtnCardSearch_Click(null, null);
            }
            else
            {
                string connStr = BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "DELETE FROM tbsir000_01 WHERE slip_seq = @seq";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@seq", seq);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    ShowToast("카드결제내역이 성공적으로 삭제되었습니다.", ColorEmerald);
                    ClearCardForm();
                    BtnCardSearch_Click(null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("카드결제내역 삭제 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // --- 과거 이력 관리 탭 초기화 및 비즈니스 로직 ---

        private void InitializePastHistoryTab()
        {
            // 상단 검색 및 설명 패널
            Panel pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 85,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            _tabPastHistoryManagement.Controls.Add(pnlTop);

            Label lblInfo = new Label
            {
                Text = "※ 특정 차트번호에 등록된 모든 마스터 이력(tbsit000_01)을 조회하여 수정하거나 불필요한 비활성(cusact='0') 이력을 삭제합니다.",
                Location = new Point(15, 12),
                Size = new Size(800, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Italic)
            };
            pnlTop.Controls.Add(lblInfo);

            Label lblChartNo = new Label { Text = "차트번호", Location = new Point(15, 45), Size = new Size(70, 20), ForeColor = ColorTextSec, Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold) };
            _txtHistoryChartNo = new TextBox { Location = new Point(90, 42), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            _txtHistoryChartNo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadHistoryData(); };
            pnlTop.Controls.Add(lblChartNo);
            pnlTop.Controls.Add(_txtHistoryChartNo);

            _btnHistorySearch = new Button
            {
                Text = "🔍 이력 조회",
                Location = new Point(220, 39),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnHistorySearch.FlatAppearance.BorderSize = 0;
            _btnHistorySearch.Click += (s, e) => LoadHistoryData();
            pnlTop.Controls.Add(_btnHistorySearch);

            // 하단 조작 패널
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 76,
                BackColor = ColorBgCard,
                Padding = new Padding(12, 10, 12, 10)
            };
            _tabPastHistoryManagement.Controls.Add(pnlBottom);

            _btnHistorySave = new Button { Text = "💾 수정사항 저장", Location = new Point(15, 18), Size = new Size(175, 40), FlatStyle = FlatStyle.Flat, BackColor = ColorEmerald, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnHistorySave.FlatAppearance.BorderSize = 0;
            _btnHistorySave.Click += BtnHistorySave_Click;
            pnlBottom.Controls.Add(_btnHistorySave);

            _btnHistoryDelete = new Button { Text = "🗑️ 선택 이력 삭제", Location = new Point(202, 18), Size = new Size(180, 40), FlatStyle = FlatStyle.Flat, BackColor = ColorAlarm, ForeColor = Color.White, Font = new Font("맑은 고딕", 9F, FontStyle.Bold) };
            _btnHistoryDelete.FlatAppearance.BorderSize = 0;
            _btnHistoryDelete.Click += BtnHistoryDelete_Click;
            pnlBottom.Controls.Add(_btnHistoryDelete);

            Action layoutHistoryButtons = delegate
            {
                int top = Math.Max(10, (pnlBottom.ClientSize.Height - 40) / 2);
                int x = 15;
                _btnHistorySave.Top = top;
                _btnHistorySave.Height = 40;
                _btnHistorySave.Width = Math.Max(175,
                    TextRenderer.MeasureText(_btnHistorySave.Text, _btnHistorySave.Font,
                        new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width + 54);
                _btnHistorySave.Left = x;
                x = _btnHistorySave.Right + 12;

                _btnHistoryDelete.Top = top;
                _btnHistoryDelete.Height = 40;
                _btnHistoryDelete.Width = Math.Max(180,
                    TextRenderer.MeasureText(_btnHistoryDelete.Text, _btnHistoryDelete.Font,
                        new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPrefix).Width + 54);
                _btnHistoryDelete.Left = x;
            };
            pnlBottom.Resize += delegate { layoutHistoryButtons(); };

            // 중앙 그리드
            _dgvHistoryMaster = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };

            _dgvHistoryMaster.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvHistoryMaster.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvHistoryMaster.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvHistoryMaster.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvHistoryMaster.DefaultCellStyle.SelectionBackColor = ColorIndigo;

            _tabPastHistoryManagement.Controls.Add(_dgvHistoryMaster);
            _dgvHistoryMaster.SendToBack();
            pnlTop.BringToFront();
            pnlBottom.BringToFront();
            layoutHistoryButtons();
        }

        public void LoadHistoryData()
        {
            string chrtno = _txtHistoryChartNo.Text.Trim();
            if (string.IsNullOrEmpty(chrtno))
            {
                MessageBox.Show("조회할 차트번호를 입력해주세요.", "차트번호 미입력", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("pat_seq", typeof(int));
            dt.Columns.Add("fam_nm", typeof(string));
            dt.Columns.Add("ADMT_TO_DT", typeof(string)); // 적용일자
            dt.Columns.Add("ADMT_FR_DT", typeof(string)); // 취득일자
            dt.Columns.Add("ins_number", typeof(string));
            dt.Columns.Add("cusact", typeof(string));

            if (_isDemo)
            {
                var list = _mockCustList.FindAll(c => c.ChrtNo == chrtno);
                int seq = 1;
                foreach (var c in list)
                {
                    string toDt = string.IsNullOrEmpty(c.HFrDt) ? (seq == 1 ? "2026-06-17" : (seq == 2 ? "2022-06-24" : "2021-08-20")) : c.HFrDt; // 적용일자
                    string frDt = string.IsNullOrEmpty(c.HToDt) ? (seq == 1 ? "2026-02-08" : (seq == 2 ? "2021-12-18" : "2021-07-20")) : c.HToDt; // 취득일자
                    string insNo = string.IsNullOrEmpty(c.InsNumber) ? (seq == 1 ? "13001531295" : (seq == 2 ? "80808084934" : "81116283481")) : c.InsNumber;
                    string famNm = string.IsNullOrEmpty(c.FamNm) ? (seq == 1 ? "이인순" : (seq == 2 ? "이인순" : "김승학")) : c.FamNm;
                    int pSeq = c.PatSeq == 0 ? seq : c.PatSeq;

                    dt.Rows.Add(pSeq, famNm, toDt, frDt, insNo, c.CusAct);
                    seq++;
                }
            }
            else
            {
                string connStr = BuildConnectionString(false);
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    try
                    {
                        conn.Open();
                        string sql = "SELECT pat_seq, fam_nm, ADMT_TO_DT, ADMT_FR_DT, ins_number, cusact FROM tbsit000_01 WITH (NOLOCK) WHERE chrtno = @chrtno ORDER BY pat_seq ASC";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@chrtno", chrtno);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }

                        // 날짜 포맷 정리 (8자리 YYYYMMDD -> YYYY-MM-DD, 적용일자 NULL/공백은 오늘 날짜 표기)
                        foreach (DataRow r in dt.Rows)
                        {
                            string toDt = r["ADMT_TO_DT"].ToString().Trim();
                            string frDt = r["ADMT_FR_DT"].ToString().Trim();
                            
                            if (string.IsNullOrEmpty(toDt))
                            {
                                toDt = DateTime.Today.ToString("yyyyMMdd");
                            }
                            
                            if (toDt.Length == 8) r["ADMT_TO_DT"] = string.Format("{0}-{1}-{2}", toDt.Substring(0, 4), toDt.Substring(4, 2), toDt.Substring(6, 2));
                            if (frDt.Length == 8) r["ADMT_FR_DT"] = string.Format("{0}-{1}-{2}", frDt.Substring(0, 4), frDt.Substring(4, 2), frDt.Substring(6, 2));
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("데이터를 로드하는 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            _dgvHistoryMaster.DataSource = dt;

            if (_dgvHistoryMaster.Columns["pat_seq"] != null)
            {
                _dgvHistoryMaster.Columns["pat_seq"].HeaderText = "순번";
                _dgvHistoryMaster.Columns["pat_seq"].Width = 80;
                _dgvHistoryMaster.Columns["pat_seq"].ReadOnly = true;
                _dgvHistoryMaster.Columns["pat_seq"].DefaultCellStyle.BackColor = ColorBgMain;
                _dgvHistoryMaster.Columns["pat_seq"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (_dgvHistoryMaster.Columns["fam_nm"] != null)
            {
                _dgvHistoryMaster.Columns["fam_nm"].HeaderText = "피보험자";
                _dgvHistoryMaster.Columns["fam_nm"].Width = 130;
                _dgvHistoryMaster.Columns["fam_nm"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (_dgvHistoryMaster.Columns["ADMT_TO_DT"] != null)
            {
                _dgvHistoryMaster.Columns["ADMT_TO_DT"].HeaderText = "적용일자";
                _dgvHistoryMaster.Columns["ADMT_TO_DT"].Width = 140;
                _dgvHistoryMaster.Columns["ADMT_TO_DT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (_dgvHistoryMaster.Columns["ADMT_FR_DT"] != null)
            {
                _dgvHistoryMaster.Columns["ADMT_FR_DT"].HeaderText = "취득일자";
                _dgvHistoryMaster.Columns["ADMT_FR_DT"].Width = 140;
                _dgvHistoryMaster.Columns["ADMT_FR_DT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (_dgvHistoryMaster.Columns["ins_number"] != null)
            {
                _dgvHistoryMaster.Columns["ins_number"].HeaderText = "보험증번호";
                _dgvHistoryMaster.Columns["ins_number"].Width = 180;
                _dgvHistoryMaster.Columns["ins_number"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (_dgvHistoryMaster.Columns["cusact"] != null)
            {
                _dgvHistoryMaster.Columns["cusact"].HeaderText = "상태";
                _dgvHistoryMaster.Columns["cusact"].Width = 80;
                _dgvHistoryMaster.Columns["cusact"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                _dgvHistoryMaster.Columns["cusact"].Visible = true;
            }

            // cusact = "1" 인 대표 활성 행 하이라이트 (연한 보라색 인디고 배경)
            foreach (DataGridViewRow row in _dgvHistoryMaster.Rows)
            {
                if (row.Cells["cusact"].Value != null && row.Cells["cusact"].Value.ToString() == "1")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(49, 46, 129); // 어두운 인디고
                    row.DefaultCellStyle.ForeColor = Color.White;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = ColorBgCard;
                    row.DefaultCellStyle.ForeColor = ColorTextMain;
                }
            }
        }

        private void BtnHistorySave_Click(object sender, EventArgs e)
        {
            string chrtno = _txtHistoryChartNo.Text.Trim();
            if (string.IsNullOrEmpty(chrtno)) return;

            _dgvHistoryMaster.EndEdit();
            DataTable dt = (DataTable)_dgvHistoryMaster.DataSource;
            if (dt == null) return;

            if (_isDemo)
            {
                var list = _mockCustList.FindAll(c => c.ChrtNo == chrtno);
                for (int i = 0; i < dt.Rows.Count && i < list.Count; i++)
                {
                    var r = dt.Rows[i];
                    list[i].FamNm = r["fam_nm"].ToString();
                    list[i].HFrDt = r["ADMT_TO_DT"].ToString(); // 적용일자
                    list[i].HToDt = r["ADMT_FR_DT"].ToString(); // 취득일자
                    list[i].InsNumber = r["ins_number"].ToString();
                    list[i].CusAct = r["cusact"].ToString();
                }
                MessageBox.Show("[데모] 수정사항이 메모리에 성공적으로 저장되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadHistoryData();
                if (_troubleshooter != null) _troubleshooter.LoadScannerGrid();
            }
            else
            {
                string connStr = BuildConnectionString(false);
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    SqlTransaction trans = null;
                    try
                    {
                        conn.Open();
                        trans = conn.BeginTransaction();
                        string sql = @"
                            UPDATE tbsit000_01
                            SET fam_nm = @fam_nm,
                                ADMT_TO_DT = @ADMT_TO_DT,
                                ADMT_FR_DT = @ADMT_FR_DT,
                                ins_number = @ins_number,
                                cusact = @cusact,
                                proc_dtime = @proc_dtime
                            WHERE chrtno = @chrtno AND pat_seq = @pat_seq";

                        string timeStr = DateTime.Now.ToString("yyyyMMddHHmmss");

                        foreach (DataRow r in dt.Rows)
                        {
                            using (SqlCommand cmd = new SqlCommand(sql, conn, trans))
                            {
                                string to = r["ADMT_TO_DT"].ToString().Replace("-", "").Trim();
                                string fr = r["ADMT_FR_DT"].ToString().Replace("-", "").Trim();

                                cmd.Parameters.AddWithValue("@fam_nm", r["fam_nm"].ToString());
                                cmd.Parameters.AddWithValue("@ADMT_TO_DT", to);
                                cmd.Parameters.AddWithValue("@ADMT_FR_DT", fr);
                                cmd.Parameters.AddWithValue("@ins_number", r["ins_number"].ToString());
                                cmd.Parameters.AddWithValue("@cusact", r["cusact"].ToString());
                                cmd.Parameters.AddWithValue("@proc_dtime", timeStr);
                                cmd.Parameters.AddWithValue("@chrtno", chrtno);
                                cmd.Parameters.AddWithValue("@pat_seq", Convert.ToInt32(r["pat_seq"]));

                                cmd.ExecuteNonQuery();
                            }
                        }

                        trans.Commit();
                        MessageBox.Show("수정사항이 데이터베이스에 성공적으로 반영되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadHistoryData();
                        if (_troubleshooter != null) _troubleshooter.LoadScannerGrid();
                    }
                    catch (Exception ex)
                    {
                        if (trans != null) trans.Rollback();
                        MessageBox.Show("저장 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnHistoryDelete_Click(object sender, EventArgs e)
        {
            string chrtno = _txtHistoryChartNo.Text.Trim();
            if (string.IsNullOrEmpty(chrtno)) return;

            if (_dgvHistoryMaster.SelectedRows.Count == 0)
            {
                MessageBox.Show("삭제할 행을 먼저 선택해주세요.", "안내", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow selRow = _dgvHistoryMaster.SelectedRows[0];
            string cusact = selRow.Cells["cusact"].Value.ToString();
            int patSeq = Convert.ToInt32(selRow.Cells["pat_seq"].Value);
            string famNm = selRow.Cells["fam_nm"].Value.ToString();

            if (cusact == "1")
            {
                MessageBox.Show("현재 활성화되어 사용 중인 대표 마스터 정보(cusact='1')는 삭제할 수 없습니다.\n먼저 활성 상태를 변경하거나 다른 레코드를 선택해 주세요.", "삭제 불가", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show(string.Format("정말로 피보험자 [{0}]의 해당 비활성 마스터 이력(순번: {1})을 삭제하시겠습니까?", famNm, patSeq), "삭제 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            if (_isDemo)
            {
                var list = _mockCustList.FindAll(c => c.ChrtNo == chrtno);
                if (patSeq > 0 && patSeq <= list.Count)
                {
                    _mockCustList.Remove(list[patSeq - 1]);
                    MessageBox.Show("[데모] 선택한 과거 마스터 이력이 메모리에서 삭제되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadHistoryData();
                    if (_troubleshooter != null) _troubleshooter.LoadScannerGrid();
                }
            }
            else
            {
                string connStr = BuildConnectionString(false);
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    try
                    {
                        conn.Open();
                        string sql = "DELETE FROM tbsit000_01 WHERE chrtno = @chrtno AND pat_seq = @pat_seq";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@chrtno", chrtno);
                            cmd.Parameters.AddWithValue("@pat_seq", patSeq);
                            int affected = cmd.ExecuteNonQuery();
                            if (affected > 0)
                            {
                                MessageBox.Show("선택한 과거 마스터 이력이 성공적으로 삭제되었습니다.", "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                LoadHistoryData();
                                if (_troubleshooter != null) _troubleshooter.LoadScannerGrid();
                            }
                            else
                            {
                                MessageBox.Show("삭제 대상 데이터를 찾지 못했습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("삭제 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public void SelectDispenseCustomerTab(TabPage tab)
        {
            _tabControl.SelectedTab = _tabDispenseCustomerManagement;
            _subTabDispenseCustomer.SelectedTab = tab;
        }

        // ==========================================
        // UI Layout & Logic - Narcotics Management
        // ==========================================

        private void InitializeNarcoticsManagementTab()
        {
            _tabNarcoticsManagement = new TabPage
            {
                Text = "💊 마약류 취급 관련",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabNarcoticsManagement);

            // Sub TabControl inside the main tab
            _subTabNarcotics = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular)
            };
            _tabNarcoticsManagement.Controls.Add(_subTabNarcotics);

            // Sub Tab 1: seq중복수정
            _tabSeqCorrection = new TabPage
            {
                Text = "🔄 seq중복수정",
                BackColor = ColorBgMain
            };
            _subTabNarcotics.TabPages.Add(_tabSeqCorrection);

            // Top Panel inside the sub tab
            Panel pnlNarcoticsTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 65,
                BackColor = ColorBgCard,
                Padding = new Padding(10)
            };
            _tabSeqCorrection.Controls.Add(pnlNarcoticsTop);

            Label lblNarcoticsTitle = new Label
            {
                Text = "💊 마약류 일련번호(INPUT_SEQ) 중복/누락 검증 및 자동 보정",
                Location = new Point(12, 10),
                Size = new Size(500, 20),
                ForeColor = ColorIndigo,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            pnlNarcoticsTop.Controls.Add(lblNarcoticsTitle);

            Label lblNarcoticsSub = new Label
            {
                Text = "※ 동일 처방 및 약품 그룹 내에서 일련번호가 꼬인 내역을 스캔하고, 1부터 N까지 순차적으로 재정렬합니다.",
                Location = new Point(12, 30),
                Size = new Size(600, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Italic)
            };
            pnlNarcoticsTop.Controls.Add(lblNarcoticsSub);

            _btnScanNarcotics = new Button
            {
                Text = "🔍 중복/누락 검사 실행",
                Location = new Point(620, 15),
                Size = new Size(150, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnScanNarcotics.FlatAppearance.BorderSize = 0;
            _btnScanNarcotics.Click += BtnScanNarcotics_Click;
            pnlNarcoticsTop.Controls.Add(_btnScanNarcotics);

            _btnFixSelectedNarcotic = new Button
            {
                Text = "⚡ 선택 건 보정 실행",
                Location = new Point(780, 15),
                Size = new Size(150, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnFixSelectedNarcotic.FlatAppearance.BorderSize = 0;
            _btnFixSelectedNarcotic.Click += BtnFixSelectedNarcotic_Click;
            pnlNarcoticsTop.Controls.Add(_btnFixSelectedNarcotic);

            _btnFixAllNarcotics = new Button
            {
                Text = "🔥 전체 일괄 보정",
                Location = new Point(940, 15),
                Size = new Size(150, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnFixAllNarcotics.FlatAppearance.BorderSize = 0;
            _btnFixAllNarcotics.Click += BtnFixAllNarcotics_Click;
            pnlNarcoticsTop.Controls.Add(_btnFixAllNarcotics);

            // SplitContainer inside the sub tab
            SplitContainer splitNarcotics = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 720,
                BackColor = ColorBorder
            };
            _tabSeqCorrection.Controls.Add(splitNarcotics);
            splitNarcotics.BringToFront();

            // Left: Grid
            _dgvNarcoticErrors = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            _dgvNarcoticErrors.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvNarcoticErrors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvNarcoticErrors.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            _dgvNarcoticErrors.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvNarcoticErrors.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvNarcoticErrors.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvNarcoticErrors.DefaultCellStyle.SelectionForeColor = Color.White;
            splitNarcotics.Panel1.Controls.Add(_dgvNarcoticErrors);

            // Right: Log Console
            _txtNarcoticsLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(15, 23, 42),
                ForeColor = Color.FromArgb(34, 197, 94),
                Font = new Font("Consolas", 9.5F),
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            splitNarcotics.Panel2.Controls.Add(_txtNarcoticsLog);

            // Sub Tab 2: 사용예정수량 수불 정정
            _tabUsageQuantity = new TabPage
            {
                Text = "📋 사용예정수량 수불 정정",
                BackColor = ColorBgMain
            };
            _subTabNarcotics.TabPages.Add(_tabUsageQuantity);

            // SplitContainer to divide into Top (Controls) and Bottom (Log Console)
            SplitContainer splitUsage = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 320,
                BackColor = ColorBorder
            };
            _tabUsageQuantity.Controls.Add(splitUsage);
            splitUsage.BringToFront();

            // SplitContainer for Left (Ghost cleanup) and Right (Canceled prescriptions) inside Panel1
            SplitContainer splitControls = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 120,
                Panel2MinSize = 120,
                SplitterDistance = 550,
                BackColor = ColorBorder
            };
            splitUsage.Panel1.Controls.Add(splitControls);
            splitControls.Resize += (s, e) => NormalizeRightPanelSplit(splitControls, 520, 480);
            NormalizeRightPanelSplit(splitControls, 520, 480);

            // Left Panel: Ghost cleanup (유령 상세대기건 정리)
            Panel pnlGhost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(15)
            };
            splitControls.Panel1.Controls.Add(pnlGhost);

            Label lblGhostTitle = new Label
            {
                Text = "👻 유령 대기 내역 일괄 정리 (PMPLUS_DUMS)",
                Location = new Point(15, 15),
                Size = new Size(400, 22),
                ForeColor = ColorIndigo,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold)
            };
            pnlGhost.Controls.Add(lblGhostTitle);

            Label lblGhostDesc = new Label
            {
                Text = "처방 마스터(TBSNM020_04)가 존재하지 않으나 상세 내역만 남아 수불에 미반영을 유발하는 데이터를 일괄 삭제합니다.",
                Location = new Point(15, 42),
                Size = new Size(500, 35),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Regular)
            };
            pnlGhost.Controls.Add(lblGhostDesc);

            _btnSearchGhostDates = new Button
            {
                Text = "🔍 대상 일자 조회",
                Location = new Point(15, 85),
                Size = new Size(160, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnSearchGhostDates.FlatAppearance.BorderSize = 0;
            _btnSearchGhostDates.Click += BtnSearchGhostDates_Click;
            pnlGhost.Controls.Add(_btnSearchGhostDates);

            _btnDeleteGhostRecords = new Button
            {
                Text = "⚡ 유령 내역 일괄 삭제",
                Location = new Point(190, 85),
                Size = new Size(180, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnDeleteGhostRecords.FlatAppearance.BorderSize = 0;
            _btnDeleteGhostRecords.Click += BtnDeleteGhostRecords_Click;
            pnlGhost.Controls.Add(_btnDeleteGhostRecords);

            Label lblGhostResultTitle = new Label
            {
                Text = "대상 처방 일자 (조회 결과):",
                Location = new Point(15, 130),
                Size = new Size(200, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            pnlGhost.Controls.Add(lblGhostResultTitle);

            _txtGhostDatesResult = new TextBox
            {
                Location = new Point(15, 155),
                Size = new Size(490, 140),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true,
                Font = new Font("Consolas", 9.5F)
            };
            pnlGhost.Controls.Add(_txtGhostDatesResult);

            // Right Panel: Canceled prescriptions (취소 처방 삭제)
            Panel pnlCanceled = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(15)
            };
            splitControls.Panel2.Controls.Add(pnlCanceled);

            Label lblCanceledTitle = new Label
            {
                Text = "❌ 취소(9) 처방 대기 내역 삭제 (PM_MAIN/PMPLUS_JOBLOG)",
                Location = new Point(15, 15),
                Size = new Size(500, 22),
                ForeColor = ColorAlarm,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold)
            };
            pnlCanceled.Controls.Add(lblCanceledTitle);

            _btnScanCanceledPrescs = new Button
            {
                Text = "🔍 취소 대기 조회",
                Location = new Point(15, 45),
                Size = new Size(130, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnScanCanceledPrescs.FlatAppearance.BorderSize = 0;
            _btnScanCanceledPrescs.Click += BtnScanCanceledPrescs_Click;
            pnlCanceled.Controls.Add(_btnScanCanceledPrescs);

            _btnDeleteSelectedCanceled = new Button
            {
                Text = "🗑️ 선택 취소 삭제",
                Location = new Point(155, 45),
                Size = new Size(130, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnDeleteSelectedCanceled.FlatAppearance.BorderSize = 0;
            _btnDeleteSelectedCanceled.Click += BtnDeleteSelectedCanceled_Click;
            pnlCanceled.Controls.Add(_btnDeleteSelectedCanceled);

            _btnDeleteAllCanceled = new Button
            {
                Text = "🗑️ 전체 취소 일괄 삭제",
                Location = new Point(295, 45),
                Size = new Size(160, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnDeleteAllCanceled.FlatAppearance.BorderSize = 0;
            _btnDeleteAllCanceled.Click += BtnDeleteAllCanceled_Click;
            pnlCanceled.Controls.Add(_btnDeleteAllCanceled);

            _dgvCanceledPrescs = new DataGridView
            {
                Location = new Point(15, 85),
                Size = new Size(500, 210),
                BackgroundColor = ColorBgMain,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 28
            };
            _dgvCanceledPrescs.ColumnHeadersDefaultCellStyle.BackColor = ColorBgCard;
            _dgvCanceledPrescs.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvCanceledPrescs.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
            _dgvCanceledPrescs.DefaultCellStyle.BackColor = ColorBgMain;
            _dgvCanceledPrescs.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvCanceledPrescs.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvCanceledPrescs.DefaultCellStyle.SelectionForeColor = Color.White;
            pnlCanceled.Controls.Add(_dgvCanceledPrescs);

            // Bottom Panel: Log Console
            _txtUsageQuantityLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(15, 23, 42),
                ForeColor = Color.FromArgb(34, 197, 94),
                Font = new Font("Consolas", 9.5F),
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            splitUsage.Panel2.Controls.Add(_txtUsageQuantityLog);
        }

        private void AppendNarcoticsLog(string msg)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((Action)(() => AppendNarcoticsLog(msg)));
            }
            else
            {
                _txtNarcoticsLog.AppendText(msg);
                _txtNarcoticsLog.SelectionStart = _txtNarcoticsLog.Text.Length;
                _txtNarcoticsLog.ScrollToCaret();
            }
        }

        private void BtnScanNarcotics_Click(object sender, EventArgs e)
        {
            AppendNarcoticsLog("============================================================\r\n");
            AppendNarcoticsLog(string.Format("▶ [{0}] 마약류 일련번호 오류 검사 시퀀스 가동 시작\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendNarcoticsLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                ScanNarcoticsMock();
            }
            else
            {
                ScanNarcoticsProduction();
            }
        }

        private void ScanNarcoticsMock()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("처방일자");
            dt.Columns.Add("처방번호");
            dt.Columns.Add("환자명");
            dt.Columns.Add("약품코드");
            dt.Columns.Add("약품명");
            dt.Columns.Add("조제박스수", typeof(int));
            dt.Columns.Add("오류유형");

            dt.Rows.Add("2026-06-16", "20260616000207", "서연석", "674900480", "스틸녹스정10밀리그램(졸피뎀타르타르산염)", 12, "순번 중복");
            dt.Rows.Add("2026-06-15", "20260615000104", "김시한", "645000160", "리보트릴정(클로나제팜)", 3, "순번 누락");
            dt.Rows.Add("2026-06-14", "20260614000088", "이연순", "657200470", "알프람정0.25밀리그램(알프라졸람)", 5, "순번 중복");
            dt.Rows.Add("2026-06-12", "20260612000311", "박보순", "642901160", "아티반정1밀리그람(로라제팜)", 4, "순번 누락");

            _dgvNarcoticErrors.DataSource = dt;

            if (_dgvNarcoticErrors.Columns.Count > 0)
            {
                _dgvNarcoticErrors.Columns[0].Width = 100;
                _dgvNarcoticErrors.Columns[1].Width = 140;
                _dgvNarcoticErrors.Columns[2].Width = 80;
                _dgvNarcoticErrors.Columns[3].Width = 100;
                _dgvNarcoticErrors.Columns[4].Width = 180;
                _dgvNarcoticErrors.Columns[5].Width = 80;
                _dgvNarcoticErrors.Columns[6].Width = 90;
            }

            AppendNarcoticsLog(string.Format("[데모] 4건의 일련번호 오류가 검출되었습니다.\r\n"));
            ShowToast("오류 검출 완료 (데모)", ColorEmerald);
        }

        private void ScanNarcoticsProduction()
        {
            if (_cmbDatabases.SelectedItem == null)
            {
                MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string connStr = BuildConnectionString(false);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
            builder.InitialCatalog = "PMPLUS_DUMS";
            string targetConnStr = builder.ConnectionString;

            this.Cursor = Cursors.WaitCursor;
            try
            {
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(targetConnStr))
                {
                    conn.Open();
                    string sql = @"
                        WITH SeqStats AS (
                            SELECT 
                                r.DRUG_SEQ,
                                r.DRUG_CODE,
                                MIN(r.ARTCNM) as DrugName,
                                COUNT(*) as RowCnt,
                                COUNT(DISTINCT r.INPUT_SEQ) as DistInputSeq,
                                MAX(CASE WHEN ISNUMERIC(r.INPUT_SEQ) = 1 THEN CAST(r.INPUT_SEQ AS INT) ELSE 0 END) as MaxInputSeq
                            FROM TBSNM020_05 r WITH (NOLOCK)
                            GROUP BY r.DRUG_SEQ, r.DRUG_CODE
                        )
                        SELECT 
                            ISNULL(m.PRES_DTIME, '') as [처방일자],
                            s.DRUG_SEQ as [처방번호],
                            ISNULL(m.PAT_NM, N'미등록') as [환자명],
                            s.DRUG_CODE as [약품코드],
                            s.DrugName as [약품명],
                            s.RowCnt as [조제박스수],
                            CASE 
                                WHEN s.RowCnt <> s.DistInputSeq THEN N'순번 중복'
                                WHEN s.RowCnt <> s.MaxInputSeq THEN N'순번 누락'
                                ELSE N'정상'
                            END as [오류유형]
                        FROM SeqStats s
                        LEFT JOIN TBSNM020_04 m WITH (NOLOCK) ON s.DRUG_SEQ = m.DRUG_SEQ
                        WHERE s.RowCnt <> s.DistInputSeq 
                           OR s.RowCnt <> s.MaxInputSeq
                        ORDER BY [처방일자] DESC, [처방번호] DESC;";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }

                _dgvNarcoticErrors.DataSource = dt;

                if (_dgvNarcoticErrors.Columns.Count > 0)
                {
                    _dgvNarcoticErrors.Columns[0].Width = 100;
                    _dgvNarcoticErrors.Columns[1].Width = 140;
                    _dgvNarcoticErrors.Columns[2].Width = 80;
                    _dgvNarcoticErrors.Columns[3].Width = 100;
                    _dgvNarcoticErrors.Columns[4].Width = 180;
                    _dgvNarcoticErrors.Columns[5].Width = 80;
                    _dgvNarcoticErrors.Columns[6].Width = 90;
                }

                AppendNarcoticsLog(string.Format("검사 완료: {0}건의 오류가 검출되었습니다.\r\n", dt.Rows.Count));
                ShowToast(string.Format("검색 완료: {0}건 조회", dt.Rows.Count), ColorEmerald);
            }
            catch (Exception ex)
            {
                AppendNarcoticsLog(string.Format("오류 검사 중 예외 발생:\r\n{0}\r\n", ex.Message));
                MessageBox.Show("마약류 일련번호 오류 검사 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void BtnFixSelectedNarcotic_Click(object sender, EventArgs e)
        {
            if (_dgvNarcoticErrors.CurrentRow == null)
            {
                MessageBox.Show("보정할 처방 내역을 목록에서 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = _dgvNarcoticErrors.CurrentRow;
            string drugSeq = row.Cells["처방번호"].Value.ToString();
            string drugCode = row.Cells["약품코드"].Value.ToString();
            string patName = row.Cells["환자명"].Value.ToString();
            string drugName = row.Cells["약품명"].Value.ToString();

            DialogResult dr = MessageBox.Show(
                string.Format("선택한 처방 건의 일련번호를 보정하시겠습니까?\n\n- 환자명: {0}\n- 처방번호: {1}\n- 약품명: {2}", patName, drugSeq, drugName),
                "일련번호 보정 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dr != DialogResult.Yes) return;

            if (_chkDemoMode.Checked)
            {
                FixNarcoticDemo(drugSeq, drugCode, patName);
            }
            else
            {
                FixNarcoticProduction(drugSeq, drugCode, patName);
            }
        }

        private void FixNarcoticDemo(string drugSeq, string drugCode, string patName)
        {
            AppendNarcoticsLog(string.Format("[데모] [{0}] 처방번호: {1}, 약품코드: {2} ({3}님 건) 보정 시작...\r\n", 
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), drugSeq, drugCode, patName));
            AppendNarcoticsLog("  - TBSNM020_05 일련번호 보정 완료 (Dynamic CTE)\r\n");
            AppendNarcoticsLog("  - TBSNM020_07 일련번호 보정 완료 (Dynamic CTE)\r\n");
            AppendNarcoticsLog(string.Format("✔ 환자 [{0}] 건 보정 완료.\r\n", patName));

            DataTable dt = _dgvNarcoticErrors.DataSource as DataTable;
            if (dt != null)
            {
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    if (dt.Rows[i]["처방번호"].ToString() == drugSeq && dt.Rows[i]["약품코드"].ToString() == drugCode)
                    {
                        dt.Rows.RemoveAt(i);
                        break;
                    }
                }
            }

            ShowToast("보정 완료 (데모)", ColorEmerald);
        }

        private void FixNarcoticProduction(string drugSeq, string drugCode, string patName)
        {
            string connStr = BuildConnectionString(false);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
            builder.InitialCatalog = "PMPLUS_DUMS";
            string targetConnStr = builder.ConnectionString;

            this.Cursor = Cursors.WaitCursor;
            try
            {
                using (SqlConnection conn = new SqlConnection(targetConnStr))
                {
                    conn.Open();
                    using (SqlTransaction trans = conn.BeginTransaction())
                    {
                        string errorMsg;
                        if (FixNarcoticSingle(conn, trans, drugSeq, drugCode, out errorMsg))
                        {
                            trans.Commit();
                            AppendNarcoticsLog(string.Format("[{0}] 처방번호: {1}, 약품코드: {2} ({3}님 건) 보정 성공.\r\n", 
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), drugSeq, drugCode, patName));
                            ShowToast("보정 성공", ColorEmerald);

                            ScanNarcoticsProduction();
                        }
                        else
                        {
                            trans.Rollback();
                            AppendNarcoticsLog(string.Format("❌ 처방번호: {0}, 약품코드: {1} 보정 실패. 원인: {2}\r\n", drugSeq, drugCode, errorMsg));
                            MessageBox.Show("보정에 실패하였습니다. 오류 로그를 확인해주십시오.\n\n오류 내용:\n" + errorMsg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendNarcoticsLog(string.Format("보정 중 예외 발생:\r\n{0}\r\n", ex.Message));
                MessageBox.Show("보정 작업 진행 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private bool FixNarcoticSingle(SqlConnection conn, SqlTransaction trans, string drugSeq, string drugCode, out string errorMsg)
        {
            errorMsg = null;
            try
            {
                string sql05 = @"
                    WITH SeqCTE AS (
                        SELECT INPUT_SEQ, ROW_NUMBER() OVER (ORDER BY IDX ASC) as NewSeq
                        FROM TBSNM020_05
                        WHERE DRUG_SEQ = @drugSeq AND DRUG_CODE = @drugCode
                    )
                    UPDATE SeqCTE SET INPUT_SEQ = CAST(NewSeq AS nvarchar(2));";

                using (SqlCommand cmd05 = new SqlCommand(sql05, conn, trans))
                {
                    cmd05.Parameters.AddWithValue("@drugSeq", drugSeq);
                    cmd05.Parameters.AddWithValue("@drugCode", drugCode);
                    cmd05.ExecuteNonQuery();
                }

                string sql07 = @"
                    WITH SeqCTE AS (
                        SELECT INPUT_SEQ, ROW_NUMBER() OVER (ORDER BY IDX ASC) as NewSeq
                        FROM TBSNM020_07
                        WHERE DRUG_SEQ = @drugSeq AND DRUG_CODE = @drugCode
                    )
                    UPDATE SeqCTE SET INPUT_SEQ = CAST(NewSeq AS nvarchar(2));";

                using (SqlCommand cmd07 = new SqlCommand(sql07, conn, trans))
                {
                    cmd07.Parameters.AddWithValue("@drugSeq", drugSeq);
                    cmd07.Parameters.AddWithValue("@drugCode", drugCode);
                    cmd07.ExecuteNonQuery();
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
                return false;
            }
        }

        private void BtnFixAllNarcotics_Click(object sender, EventArgs e)
        {
            if (_dgvNarcoticErrors.Rows.Count == 0)
            {
                MessageBox.Show("보정할 오류 내역이 없습니다. 먼저 검사를 실행해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int totalCount = _dgvNarcoticErrors.Rows.Count;
            DialogResult dr = MessageBox.Show(
                string.Format("검출된 전체 {0}건의 일련번호 오류에 대해 일괄 보정을 실행하시겠습니까?\n\n이 작업은 각 건별로 개별 트랜잭션을 적용하여 순차적으로 보정합니다.", totalCount),
                "전체 일괄 보정 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            _btnScanNarcotics.Enabled = false;
            _btnFixSelectedNarcotic.Enabled = false;
            _btnFixAllNarcotics.Enabled = false;

            if (_chkDemoMode.Checked)
            {
                FixAllNarcoticsDemo();
            }
            else
            {
                FixAllNarcoticsProduction();
            }
        }

        private void FixAllNarcoticsDemo()
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    DataTable dt = null;
                    this.Invoke((Action)(() => {
                        dt = _dgvNarcoticErrors.DataSource as DataTable;
                    }));

                    if (dt == null) return;

                    List<DataRow> rowsToFix = new List<DataRow>();
                    foreach (DataRow row in dt.Rows)
                    {
                        rowsToFix.Add(row);
                    }

                    int successCount = 0;
                    foreach (var row in rowsToFix)
                    {
                        string drugSeq = row["처방번호"].ToString();
                        string drugCode = row["약품코드"].ToString();
                        string patName = row["환자명"].ToString();

                        AppendNarcoticsLog(string.Format("[데모] [{0}] 처방번호: {1}, 약품코드: {2} ({3}님 건) 보정 중...\r\n", 
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), drugSeq, drugCode, patName));

                        System.Threading.Thread.Sleep(200);
                        successCount++;
                    }

                    this.Invoke((Action)(() => {
                        dt.Clear();
                        AppendNarcoticsLog(string.Format("\r\n============================================================\r\n"));
                        AppendNarcoticsLog(string.Format("✔ [데모] 전체 일괄 보정 작업 완료 (총 {0}건 완료)\r\n", successCount));
                        AppendNarcoticsLog(string.Format("============================================================\r\n"));
                        ShowToast("일괄 보정 완료 (데모)", ColorEmerald);
                    }));
                }
                catch (Exception ex)
                {
                    AppendNarcoticsLog("일괄 보정 중 오류 발생: " + ex.Message + "\r\n");
                }
                finally
                {
                    this.Invoke((Action)(() => {
                        _btnScanNarcotics.Enabled = true;
                        _btnFixSelectedNarcotic.Enabled = true;
                        _btnFixAllNarcotics.Enabled = true;
                    }));
                }
            });
        }

        private void FixAllNarcoticsProduction()
        {
            string connStr = BuildConnectionString(false);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(connStr);
            builder.InitialCatalog = "PMPLUS_DUMS";
            string targetConnStr = builder.ConnectionString;

            List<Tuple<string, string, string>> itemsToFix = new List<Tuple<string, string, string>>();
            foreach (DataGridViewRow row in _dgvNarcoticErrors.Rows)
            {
                if (row.IsNewRow) continue;
                string drugSeq = row.Cells["처방번호"].Value.ToString();
                string drugCode = row.Cells["약품코드"].Value.ToString();
                string patName = row.Cells["환자명"].Value.ToString();
                itemsToFix.Add(new Tuple<string, string, string>(drugSeq, drugCode, patName));
            }

            this.Cursor = Cursors.WaitCursor;
            System.Threading.Tasks.Task.Run(() =>
            {
                int successCount = 0;
                int failCount = 0;

                try
                {
                    using (SqlConnection conn = new SqlConnection(targetConnStr))
                    {
                        conn.Open();
                        for (int i = 0; i < itemsToFix.Count; i++)
                        {
                            var item = itemsToFix[i];
                            string drugSeq = item.Item1;
                            string drugCode = item.Item2;
                            string patName = item.Item3;

                            AppendNarcoticsLog(string.Format("[{0}/{1}] 처방번호: {2}, 약품코드: {3} ({4}님 건) 보정 시작...\r\n", 
                                i + 1, itemsToFix.Count, drugSeq, drugCode, patName));

                            using (SqlTransaction trans = conn.BeginTransaction())
                            {
                                string errorMsg;
                                if (FixNarcoticSingle(conn, trans, drugSeq, drugCode, out errorMsg))
                                {
                                    trans.Commit();
                                    successCount++;
                                    AppendNarcoticsLog(string.Format(" ➔ ✔ 보정 성공\r\n"));
                                }
                                else
                                {
                                    trans.Rollback();
                                    failCount++;
                                    AppendNarcoticsLog(string.Format(" ➔ ❌ 보정 실패: {0}\r\n", errorMsg));
                                }
                            }
                        }
                    }

                    this.Invoke((Action)(() => {
                        AppendNarcoticsLog(string.Format("\r\n============================================================\r\n"));
                        AppendNarcoticsLog(string.Format("▶ [{0}] 전체 일괄 보정 작업 완료\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                        AppendNarcoticsLog(string.Format("- 성공: {0}건, 실패: {1}건\r\n", successCount, failCount));
                        AppendNarcoticsLog(string.Format("============================================================\r\n"));

                        ScanNarcoticsProduction();
                    }));
                }
                catch (Exception ex)
                {
                    AppendNarcoticsLog(string.Format("일괄 보정 처리 중 중대한 예외 발생:\r\n{0}\r\n", ex.Message));
                    this.Invoke((Action)(() => {
                        MessageBox.Show("일괄 보정 처리 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    this.Invoke((Action)(() => {
                        this.Cursor = Cursors.Default;
                        _btnScanNarcotics.Enabled = true;
                        _btnFixSelectedNarcotic.Enabled = true;
                        _btnFixAllNarcotics.Enabled = true;
                    }));
                }
            });
        }

        // ====================================================
        // UI Layout & Logic - Narcotics Usage Quantity Cleanup
        // ====================================================

        private void AppendUsageQuantityLog(string msg)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((Action)(() => AppendUsageQuantityLog(msg)));
            }
            else
            {
                _txtUsageQuantityLog.AppendText(msg);
                _txtUsageQuantityLog.SelectionStart = _txtUsageQuantityLog.Text.Length;
                _txtUsageQuantityLog.ScrollToCaret();
            }
        }

        private void BtnSearchGhostDates_Click(object sender, EventArgs e)
        {
            AppendUsageQuantityLog("============================================================\r\n");
            AppendUsageQuantityLog(string.Format("▶ [{0}] 유령 상세 대기 내역 일자 조회 시작...\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendUsageQuantityLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                _txtGhostDatesResult.Text = "2026-06-10\r\n2026-06-11\r\n(데모 모드 가상 결과)";
                AppendUsageQuantityLog("[데모] 유령 상세 내역 대상 일자가 조회되었습니다.\r\n");
                ShowToast("대상 조회 완료 (데모)", ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    List<string> dates = new List<string>();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT DISTINCT LEFT(a.drug_Seq, 8) 
                            FROM pmplus_dums..tbsnm020_05 a WITH (NOLOCK)
                            WHERE a.REPORT_GUBUN = '2' AND a.SEND_GUBUN = '0'
                            AND NOT EXISTS (
                                SELECT 1 FROM pmplus_dums..tbsnm020_04 b WITH (NOLOCK)
                                WHERE b.REPORT_GUBUN = '2' AND b.SEND_GUBUN = '0'
                                AND a.drug_Seq = b.drug_Seq
                            );";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    dates.Add(reader.GetValue(0).ToString());
                                }
                            }
                        }
                    }

                    if (dates.Count > 0)
                    {
                        _txtGhostDatesResult.Text = string.Join("\r\n", dates.ToArray());
                        AppendUsageQuantityLog(string.Format("조회 완료: {0}개의 일자가 검출되었습니다.\r\n", dates.Count));
                        ShowToast(string.Format("조회 완료: {0}건", dates.Count), ColorEmerald);
                    }
                    else
                    {
                        _txtGhostDatesResult.Text = "(검출된 대상 일자 없음)";
                        AppendUsageQuantityLog("검사 완료: 보정 대상 유령 내역이 존재하지 않습니다.\r\n");
                        ShowToast("대상 일자 없음", ColorEmerald);
                    }
                }
                catch (Exception ex)
                {
                    AppendUsageQuantityLog("오류 발생: " + ex.Message + "\r\n");
                    MessageBox.Show("유령 내역 조회 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnDeleteGhostRecords_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(
                "👻 유령 대기 내역(마스터가 없는 상세내역)을 일괄 삭제하시겠습니까?\n\n이 작업은 데이터베이스에서 물리적인 삭제(DELETE) 작업을 실행합니다.",
                "유령 대기 내역 삭제 경고",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            AppendUsageQuantityLog("============================================================\r\n");
            AppendUsageQuantityLog(string.Format("▶ [{0}] 유령 대기 내역 일괄 삭제 작업 개시\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendUsageQuantityLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                AppendUsageQuantityLog("[데모] 1. pmplus_dums..tbsnm020_06 (로그 마스터) 정리 완료.\r\n");
                AppendUsageQuantityLog("[데모] 2. pmplus_dums..tbsnm020_07 (로그 상세) 정리 완료.\r\n");
                AppendUsageQuantityLog("[데모] 3. pmplus_dums..tbsnm020_05 (상세 내역) 정리 완료.\r\n");
                AppendUsageQuantityLog("✔ [데모] 유령 대기 내역 일괄 삭제 성공.\r\n");
                _txtGhostDatesResult.Text = "";
                ShowToast("삭제 성공 (데모)", ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                // Query 2: Delete tbsnm020_06
                                string sql06 = @"
                                    DELETE t1 
                                    FROM pmplus_dums..tbsnm020_06 t1, 
                                         (SELECT a.drug_Seq 
                                          FROM pmplus_dums..tbsnm020_05 a WITH (NOLOCK)
                                          WHERE a.REPORT_GUBUN = '2' AND a.SEND_GUBUN = '0'
                                          AND NOT EXISTS (
                                              SELECT 1 FROM pmplus_dums..tbsnm020_04 b WITH (NOLOCK)
                                              WHERE b.REPORT_GUBUN = '2' AND b.SEND_GUBUN = '0'
                                              AND a.drug_Seq = b.drug_Seq
                                          )) t2
                                    WHERE t1.drug_Seq = t2.drug_Seq
                                    AND t1.REPORT_GUBUN = '2' AND t1.SEND_GUBUN = '0';";

                                int affected06 = 0;
                                using (SqlCommand cmd = new SqlCommand(sql06, conn, trans))
                                {
                                    affected06 = cmd.ExecuteNonQuery();
                                }
                                AppendUsageQuantityLog(string.Format(" ➔ tbsnm020_06 (로그 마스터) 삭제 완료: {0}행\r\n", affected06));

                                // Query 3: Delete tbsnm020_07
                                string sql07 = @"
                                    DELETE t1 
                                    FROM pmplus_dums..tbsnm020_07 t1, 
                                         (SELECT a.drug_Seq 
                                          FROM pmplus_dums..tbsnm020_05 a WITH (NOLOCK)
                                          WHERE a.REPORT_GUBUN = '2' AND a.SEND_GUBUN = '0'
                                          AND NOT EXISTS (
                                              SELECT 1 FROM pmplus_dums..tbsnm020_04 b WITH (NOLOCK)
                                              WHERE b.REPORT_GUBUN = '2' AND b.SEND_GUBUN = '0'
                                              AND a.drug_Seq = b.drug_Seq
                                          )) t2
                                    WHERE t1.drug_Seq = t2.drug_Seq
                                    AND t1.REPORT_GUBUN = '2' AND t1.SEND_GUBUN = '0';";

                                int affected07 = 0;
                                using (SqlCommand cmd = new SqlCommand(sql07, conn, trans))
                                {
                                    affected07 = cmd.ExecuteNonQuery();
                                }
                                AppendUsageQuantityLog(string.Format(" ➔ tbsnm020_07 (로그 상세) 삭제 완료: {0}행\r\n", affected07));

                                // Query 4: Delete tbsnm020_05
                                string sql05 = @"
                                    DELETE a 
                                    FROM pmplus_dums..tbsnm020_05 a
                                    WHERE a.REPORT_GUBUN = '2' AND a.SEND_GUBUN = '0'
                                    AND NOT EXISTS (
                                        SELECT 1 FROM pmplus_dums..tbsnm020_04 b WITH (NOLOCK)
                                        WHERE b.REPORT_GUBUN = '2' AND b.SEND_GUBUN = '0'
                                        AND a.drug_Seq = b.drug_Seq
                                    );";

                                int affected05 = 0;
                                using (SqlCommand cmd = new SqlCommand(sql05, conn, trans))
                                {
                                    affected05 = cmd.ExecuteNonQuery();
                                }
                                AppendUsageQuantityLog(string.Format(" ➔ tbsnm020_05 (상세 내역) 삭제 완료: {0}행\r\n", affected05));

                                trans.Commit();
                                AppendUsageQuantityLog("✔ 유령 대기 내역 일괄 정리 성공 및 커밋 완료.\r\n");
                                _txtGhostDatesResult.Text = "";
                                ShowToast("정리 성공", ColorEmerald);
                            }
                            catch (Exception ex)
                            {
                                trans.Rollback();
                                AppendUsageQuantityLog("❌ 삭제 중 에러 발생 (롤백함): " + ex.Message + "\r\n");
                                MessageBox.Show("삭제 작업 도중 오류가 발생하여 모든 변경 사항이 롤백되었습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendUsageQuantityLog("연결 실패: " + ex.Message + "\r\n");
                    MessageBox.Show("데이터베이스 연결 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnScanCanceledPrescs_Click(object sender, EventArgs e)
        {
            AppendUsageQuantityLog("============================================================\r\n");
            AppendUsageQuantityLog(string.Format("▶ [{0}] 취소(9) 처방 대기 내역 스캔...\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendUsageQuantityLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("처방번호");
                dt.Columns.Add("환자명");
                dt.Columns.Add("차트번호");
                dt.Columns.Add("주민번호");
                dt.Columns.Add("교부번호");
                dt.Columns.Add("상태");

                dt.Rows.Add("20230925000043", "서연석", "0000184791", "550505-1xxxxxx", "2023092500012", "9:취소(삭제)");
                dt.Rows.Add("20230925000016", "김시한", "0000138658", "600606-1xxxxxx", "2023092500008", "9:취소(삭제)");
                dt.Rows.Add("20220610000136", "이연순", "0100028355", "700707-2xxxxxx", "2022061000030", "9:취소(삭제)");

                _dgvCanceledPrescs.DataSource = dt;
                
                if (_dgvCanceledPrescs.Columns.Count > 0)
                {
                    _dgvCanceledPrescs.Columns[0].Width = 110;
                    _dgvCanceledPrescs.Columns[1].Width = 80;
                    _dgvCanceledPrescs.Columns[2].Width = 90;
                    _dgvCanceledPrescs.Columns[3].Width = 100;
                    _dgvCanceledPrescs.Columns[4].Width = 110;
                    _dgvCanceledPrescs.Columns[5].Width = 80;
                }

                AppendUsageQuantityLog("[데모] 3건의 취소(9) 처방 대기 건이 검출되었습니다.\r\n");
                ShowToast("조회 완료 (데모)", ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    DataTable dt = new DataTable();
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = @"
                            SELECT drug_seq as [처방번호], 
                                   pat_nm as [환자명], 
                                   chrtno as [차트번호], 
                                   pat_jumin_no as [주민번호], 
                                   mprsc_Grant_no as [교부번호], 
                                   PRES_PRGRS_STATE as [상태]
                            FROM PM_MAIN..TBSID040_03 WITH (NOLOCK)
                            WHERE PRES_PRGRS_STATE = '9'
                            ORDER BY DRUG_SEQ DESC;";

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }

                    _dgvCanceledPrescs.DataSource = dt;

                    if (_dgvCanceledPrescs.Columns.Count > 0)
                    {
                        _dgvCanceledPrescs.Columns[0].Width = 110;
                        _dgvCanceledPrescs.Columns[1].Width = 80;
                        _dgvCanceledPrescs.Columns[2].Width = 90;
                        _dgvCanceledPrescs.Columns[3].Width = 100;
                        _dgvCanceledPrescs.Columns[4].Width = 110;
                        _dgvCanceledPrescs.Columns[5].Width = 80;
                    }

                    AppendUsageQuantityLog(string.Format("조회 완료: {0}건의 취소 처방 대기가 검출되었습니다.\r\n", dt.Rows.Count));
                    ShowToast(string.Format("조회 완료: {0}건", dt.Rows.Count), ColorEmerald);
                }
                catch (Exception ex)
                {
                    AppendUsageQuantityLog("오류 발생: " + ex.Message + "\r\n");
                    MessageBox.Show("취소 처방 조회 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnDeleteSelectedCanceled_Click(object sender, EventArgs e)
        {
            if (_dgvCanceledPrescs.CurrentRow == null)
            {
                MessageBox.Show("삭제할 취소 처방 건을 목록에서 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataGridViewRow row = _dgvCanceledPrescs.CurrentRow;
            string drugSeq = row.Cells["처방번호"].Value.ToString();
            string patNm = row.Cells["환자명"].Value.ToString();

            DialogResult dr = MessageBox.Show(
                string.Format("선택한 취소(9) 처방 건을 영구 삭제하시겠습니까?\n\n- 환자명: {0}\n- 처방번호: {1}\n\n※ 주의: PM_MAIN 및 PMPLUS_JOBLOG에서 완전 삭제됩니다.", patNm, drugSeq),
                "취소 처방 선택 삭제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            AppendUsageQuantityLog("============================================================\r\n");
            AppendUsageQuantityLog(string.Format("▶ [{0}] 선택 취소 처방 삭제 개시 (처방번호: {1})\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), drugSeq));
            AppendUsageQuantityLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                AppendUsageQuantityLog(string.Format("[데모] 처방번호 {0} 삭제 완료.\r\n", drugSeq));
                
                DataTable dt = _dgvCanceledPrescs.DataSource as DataTable;
                if (dt != null)
                {
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        if (dt.Rows[i]["처방번호"].ToString() == drugSeq)
                        {
                            dt.Rows.RemoveAt(i);
                            break;
                        }
                    }
                }
                ShowToast("선택 삭제 완료 (데모)", ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                int aff3 = 0, aff4 = 0, aff5 = 0, affLog = 0;

                                using (SqlCommand cmd = new SqlCommand("DELETE FROM PM_MAIN..TBSID040_03 WHERE drug_seq = @drugSeq", conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@drugSeq", drugSeq);
                                    aff3 = cmd.ExecuteNonQuery();
                                }
                                using (SqlCommand cmd = new SqlCommand("DELETE FROM PM_MAIN..TBSID040_04 WHERE drug_seq = @drugSeq", conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@drugSeq", drugSeq);
                                    aff4 = cmd.ExecuteNonQuery();
                                }
                                using (SqlCommand cmd = new SqlCommand("DELETE FROM PM_MAIN..TBSID040_05 WHERE drug_seq = @drugSeq", conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@drugSeq", drugSeq);
                                    aff5 = cmd.ExecuteNonQuery();
                                }
                                using (SqlCommand cmd = new SqlCommand("DELETE FROM PMPLUS_JOBLOG..PM_PRES_LOG WHERE PRESERIAL = @drugSeq", conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@drugSeq", drugSeq);
                                    affLog = cmd.ExecuteNonQuery();
                                }

                                trans.Commit();
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_03 삭제: {0}행\r\n", aff3));
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_04 삭제: {0}행\r\n", aff4));
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_05 삭제: {0}행\r\n", aff5));
                                AppendUsageQuantityLog(string.Format(" ➔ PM_PRES_LOG 삭제: {0}행\r\n", affLog));
                                AppendUsageQuantityLog("✔ 선택한 취소 처방 데이터 삭제 성공 및 커밋 완료.\r\n");
                                ShowToast("삭제 완료", ColorEmerald);
                                
                                BtnScanCanceledPrescs_Click(null, null);
                            }
                            catch (Exception ex)
                            {
                                trans.Rollback();
                                AppendUsageQuantityLog("❌ 삭제 실패 (롤백): " + ex.Message + "\r\n");
                                MessageBox.Show("삭제 작업 도중 오류가 발생하여 모든 변경 사항이 롤백되었습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendUsageQuantityLog("연결 실패: " + ex.Message + "\r\n");
                    MessageBox.Show("데이터베이스 연결 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void BtnDeleteAllCanceled_Click(object sender, EventArgs e)
        {
            if (_dgvCanceledPrescs.Rows.Count == 0)
            {
                MessageBox.Show("삭제할 취소 대기 내역이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int count = _dgvCanceledPrescs.Rows.Count;
            DialogResult dr = MessageBox.Show(
                string.Format("검출된 전체 {0}건의 취소(9) 처방 대기 내역을 일괄 영구 삭제하시겠습니까?\n\n※ 주의: 복구가 불가능하며 PM_MAIN 및 PMPLUS_JOBLOG의 관련 데이터가 모두 일괄 삭제됩니다.", count),
                "취소 처방 일괄 삭제 경고",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Stop
            );

            if (dr != DialogResult.Yes) return;

            AppendUsageQuantityLog("============================================================\r\n");
            AppendUsageQuantityLog(string.Format("▶ [{0}] 전체 취소 처방 일괄 삭제 작업 개시\r\n", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            AppendUsageQuantityLog("============================================================\r\n");

            if (_chkDemoMode.Checked)
            {
                AppendUsageQuantityLog(string.Format("[데모] 전체 {0}건 일괄 삭제 완료.\r\n", count));
                DataTable dt = _dgvCanceledPrescs.DataSource as DataTable;
                if (dt != null)
                {
                    dt.Clear();
                }
                ShowToast("일괄 삭제 완료 (데모)", ColorEmerald);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null)
                {
                    MessageBox.Show("데이터베이스를 먼저 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                int aff3 = 0, aff4 = 0, aff5 = 0, affLog = 0;

                                // 1. Delete from PM_PRES_LOG
                                string sqlLog = @"
                                    DELETE FROM PMPLUS_JOBLOG..PM_PRES_LOG 
                                    WHERE PRESERIAL IN (SELECT drug_seq FROM PM_MAIN..TBSID040_03 WITH (NOLOCK) WHERE PRES_PRGRS_STATE = '9');";
                                using (SqlCommand cmd = new SqlCommand(sqlLog, conn, trans))
                                {
                                    affLog = cmd.ExecuteNonQuery();
                                }

                                // 2. Delete from TBSID040_05
                                string sql5 = @"
                                    DELETE FROM PM_MAIN..TBSID040_05 
                                    WHERE drug_seq IN (SELECT drug_seq FROM PM_MAIN..TBSID040_03 WITH (NOLOCK) WHERE PRES_PRGRS_STATE = '9');";
                                using (SqlCommand cmd = new SqlCommand(sql5, conn, trans))
                                {
                                    aff5 = cmd.ExecuteNonQuery();
                                }

                                // 3. Delete from TBSID040_04
                                string sql4 = @"
                                    DELETE FROM PM_MAIN..TBSID040_04 
                                    WHERE drug_seq IN (SELECT drug_seq FROM PM_MAIN..TBSID040_03 WITH (NOLOCK) WHERE PRES_PRGRS_STATE = '9');";
                                using (SqlCommand cmd = new SqlCommand(sql4, conn, trans))
                                {
                                    aff4 = cmd.ExecuteNonQuery();
                                }

                                // 4. Delete from TBSID040_03
                                string sql3 = "DELETE FROM PM_MAIN..TBSID040_03 WHERE PRES_PRGRS_STATE = '9';";
                                using (SqlCommand cmd = new SqlCommand(sql3, conn, trans))
                                {
                                    aff3 = cmd.ExecuteNonQuery();
                                }

                                trans.Commit();
                                AppendUsageQuantityLog(string.Format(" ➔ PM_PRES_LOG 삭제: {0}행\r\n", affLog));
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_05 삭제: {0}행\r\n", aff5));
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_04 삭제: {0}행\r\n", aff4));
                                AppendUsageQuantityLog(string.Format(" ➔ TBSID040_03 삭제: {0}행\r\n", aff3));
                                AppendUsageQuantityLog("✔ 전체 취소 처방 데이터 일괄 삭제 성공 및 커밋 완료.\r\n");
                                ShowToast("일괄 삭제 성공", ColorEmerald);
                                
                                BtnScanCanceledPrescs_Click(null, null);
                            }
                            catch (Exception ex)
                            {
                                trans.Rollback();
                                AppendUsageQuantityLog("❌ 삭제 실패 (롤백): " + ex.Message + "\r\n");
                                MessageBox.Show("삭제 작업 도중 오류가 발생하여 모든 변경 사항이 롤백되었습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendUsageQuantityLog("연결 실패: " + ex.Message + "\r\n");
                    MessageBox.Show("데이터베이스 연결 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        // ====================================================
        // UI Layout & Logic - DB Map & SQL Query Runner
        // ====================================================

        private void InitializeQueryRunnerTab()
        {
            _tabQueryRunner = new TabPage
            {
                Text = "\u26A1 \uCFFC\uB9AC \uC2E4\uD589\uAE30",
                BackColor = ColorBgMain
            };
            _tabControl.TabPages.Add(_tabQueryRunner);

            SplitContainer splitQueryMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 220,
                BackColor = ColorBorder
            };
            _tabQueryRunner.Controls.Add(splitQueryMain);

            Panel pnlQueryTop = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            splitQueryMain.Panel1.Controls.Add(pnlQueryTop);

            Label lblQueryDb = new Label
            {
                Text = "\uB300\uC0C1 \uB370\uC774\uD130\uBCA0\uC774\uC2A4",
                Location = new Point(12, 12),
                Size = new Size(110, 20),
                ForeColor = ColorTextSec,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9F, FontStyle.Bold)
            };
            _cmbQueryDbSelector = new ComboBox
            {
                Location = new Point(130, 10),
                Size = new Size(180, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain
            };
            pnlQueryTop.Controls.Add(lblQueryDb);
            pnlQueryTop.Controls.Add(_cmbQueryDbSelector);

            Label lblQueryInfo = new Label
            {
                Text = "\u203B \uC8FC\uC758: SELECT \uC774\uC678\uC758 \uBCC0\uACBD \uCF7C\uB9AC \uC2E4\uD589 \uC2DC \uC548\uC804 \uC7AC\uD655\uC778 \uACBD\uACE0\uCC3D\uC774 \uC791\uB3D9\uD569\uB2C8\uB2E4. \uC885\uB8CC \uC2DC \uC138\uBBF8\uCF5C\uB860(;) \uD544\uC218.",
                Location = new Point(330, 13),
                Size = new Size(600, 20),
                ForeColor = ColorAlarm,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 8.5F, FontStyle.Italic)
            };
            pnlQueryTop.Controls.Add(lblQueryInfo);

            Panel pnlQueryRight = new Panel
            {
                Dock = DockStyle.Right,
                Width = 120,
                BackColor = ColorBgCard,
                Padding = new Padding(6, 45, 6, 6)
            };
            pnlQueryTop.Controls.Add(pnlQueryRight);

            _btnExecuteQuery = new Button
            {
                Text = "\u26a1 SQL\n\uc2e4\ud589",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 11F, FontStyle.Bold)
            };
            _btnExecuteQuery.FlatAppearance.BorderSize = 0;
            _btnExecuteQuery.Click += BtnExecuteQuery_Click;
            pnlQueryRight.Controls.Add(_btnExecuteQuery);

            _txtQueryInput = new TextBox
            {
                Location = new Point(12, 45),
                Size = new Size(pnlQueryTop.Width - 140, 130),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(15, 23, 42),
                ForeColor = Color.White,
                Font = new Font("Consolas", 10F),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _txtQueryInput.Text = "SELECT * FROM PM_MAIN..TBSID040_03 WHERE PRES_PRGRS_STATE = '9';";
            pnlQueryTop.Controls.Add(_txtQueryInput);

            Panel pnlQueryBottom = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgCard,
                Padding = new Padding(12)
            };
            splitQueryMain.Panel2.Controls.Add(pnlQueryBottom);

            Label lblQueryResultTitle = new Label
            {
                Text = "\uD83D\uDCCA \uC2E4\uD589 \uACB0\uACFC",
                Location = new Point(12, 5),
                Size = new Size(100, 20),
                ForeColor = ColorIndigo,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9.5F, FontStyle.Bold)
            };
            pnlQueryBottom.Controls.Add(lblQueryResultTitle);

            _lblQueryStatus = new Label
            {
                Text = "\uC900\uBE44 \uC644\uB8CC (SQL\uC744 \uC785\uB825\uD558\uACE0 \uC2E4\uD589 \uBC84\uD2BC\uC744 \uB204\uB974\uC138\uC694.)",
                Location = new Point(130, 5),
                Size = new Size(800, 20),
                ForeColor = ColorTextSec,
                Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9F)
            };
            pnlQueryBottom.Controls.Add(_lblQueryStatus);

            _dgvQueryResult = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgMain,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 28
            };
            _dgvQueryResult.ColumnHeadersDefaultCellStyle.BackColor = ColorBgCard;
            _dgvQueryResult.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvQueryResult.ColumnHeadersDefaultCellStyle.Font = new Font("\uB9D1\uC740 \uACE0\uB515", 9F, FontStyle.Bold);
            _dgvQueryResult.DefaultCellStyle.BackColor = ColorBgMain;
            _dgvQueryResult.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvQueryResult.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            _dgvQueryResult.DefaultCellStyle.SelectionForeColor = Color.White;
            pnlQueryBottom.Controls.Add(_dgvQueryResult);
        }
        private void LoadQueryRunnerDatabases()
        {
            if (_cmbQueryDbSelector == null) return;

            _cmbQueryDbSelector.Items.Clear();

            if (_isDemo)
            {
                string[] dbs = { "PM_MAIN", "PMPLUS_DUMS", "PMPLUS_IMAGE", "PMPLUS_JOBLOG" };
                foreach (var db in dbs)
                {
                    _cmbQueryDbSelector.Items.Add(db);
                }
                _cmbQueryDbSelector.SelectedIndex = 0;
            }
            else
            {
                try
                {
                    string connStr = BuildConnectionString(false);
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand("SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name", conn))
                        {
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string dbName = reader.GetString(0);
                                    _cmbQueryDbSelector.Items.Add(dbName);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    string[] dbs = { "PM_MAIN", "PMPLUS_DUMS", "PMPLUS_IMAGE", "PMPLUS_JOBLOG" };
                    foreach (var db in dbs)
                    {
                        _cmbQueryDbSelector.Items.Add(db);
                    }
                }

                if (_cmbQueryDbSelector.Items.Count > 0)
                {
                    string currentDb = _cmbDatabases != null && _cmbDatabases.SelectedItem != null ? _cmbDatabases.SelectedItem.ToString() : "PM_MAIN";
                    int idx = _cmbQueryDbSelector.FindStringExact(currentDb);
                    _cmbQueryDbSelector.SelectedIndex = idx >= 0 ? idx : 0;
                }
            }
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == _tabQueryRunner)
            {
                if (_cmbQueryDbSelector != null && _cmbQueryDbSelector.Items.Count == 0)
                {
                    LoadQueryRunnerDatabases();
                }
            }
            if (_tabControl.SelectedTab == _tabDataManagement)
            {
                BeginInvoke((Action)(() =>
                {
                    if (_splitUser  != null && _splitUser.Visible)  try { _splitUser.SplitterDistance  = _distUser;  } catch {}
                    if (_splitCard  != null && _splitCard.Visible)  try { _splitCard.SplitterDistance  = _distCard;  } catch {}
                    if (_splitLabel != null && _splitLabel.Visible) try { _splitLabel.SplitterDistance = _distLabel; } catch {}
                    if (_splitRx    != null && _splitRx.Visible)    NormalizeRightPanelSplit(_splitRx, ref _distRx, 340, 360);
                }));
            }
        }

        private void BtnExecuteQuery_Click(object sender, EventArgs e)
        {
            string sql = _txtQueryInput.Text.Trim();
            if (string.IsNullOrEmpty(sql))
            {
                MessageBox.Show("실행할 SQL 쿼리를 입력해 주십시오.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string targetDb = _cmbQueryDbSelector.SelectedItem != null ? _cmbQueryDbSelector.SelectedItem.ToString() : "PM_MAIN";

            // Safety check for DDL or DML modifying queries
            string lowerSql = sql.ToLower();
            string[] unsafeKeywords = { "update", "delete", "insert", "drop", "truncate", "alter", "create" };
            bool isUnsafe = false;
            foreach (var kw in unsafeKeywords)
            {
                if (lowerSql.Contains(kw))
                {
                    isUnsafe = true;
                    break;
                }
            }

            if (isUnsafe)
            {
                DialogResult dr = MessageBox.Show(
                    "⚠️ [데이터 변경 위험 경고]\n\n입력하신 쿼리에 UPDATE, DELETE, DROP 등 데이터를 물리적으로 변경하거나 구조를 해치는 명령어가 포함되어 있습니다.\n\n이 쿼리를 실제로 실행하시겠습니까?",
                    "위험 쿼리 경고 및 재확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Stop
                );
                if (dr != DialogResult.Yes) return;
            }

            _lblQueryStatus.ForeColor = ColorTextMain;
            _lblQueryStatus.Text = "쿼리 실행 중...";

            DataTable dt = new DataTable();

            if (_isDemo)
            {
                // Demo Mode Simulation output
                System.Threading.Thread.Sleep(300);
                if (lowerSql.Contains("tbsid040_03"))
                {
                    dt.Columns.Add("drug_seq");
                    dt.Columns.Add("pres_dtime");
                    dt.Columns.Add("pat_nm");
                    dt.Columns.Add("chrtno");
                    dt.Columns.Add("PRES_PRGRS_STATE");

                    dt.Rows.Add("20230925000043", "2023-09-25 10:20:10", "서연석", "0000184791", "9");
                    dt.Rows.Add("20230925000016", "2023-09-25 09:15:33", "김시한", "0000138658", "9");
                    dt.Rows.Add("20220610000136", "2022-06-10 14:05:40", "이연순", "0100028355", "9");
                }
                else if (lowerSql.Contains("tbsnm020_05"))
                {
                    dt.Columns.Add("drug_seq");
                    dt.Columns.Add("drug_code");
                    dt.Columns.Add("INPUT_SEQ");
                    dt.Columns.Add("IDX");

                    dt.Rows.Add("20230925000043", "8806446011701", "1", "110492");
                    dt.Rows.Add("20230925000043", "8806418002409", "2", "110493");
                }
                else
                {
                    dt.Columns.Add("Status");
                    dt.Columns.Add("Message");
                    dt.Rows.Add("Success", "데모 모드 쿼리가 성공적으로 모의 수행되었습니다.");
                }
                _dgvQueryResult.DataSource = dt;
                _lblQueryStatus.Text = string.Format("실행 완료 (가상): {0}행 반환됨.", dt.Rows.Count);
            }
            else
            {
                if (_cmbDatabases.SelectedItem == null || _cmbDatabases.Items.Count == 0)
                {
                    _lblQueryStatus.ForeColor = ColorAlarm;
                    _lblQueryStatus.Text = "❌ 실행 실패: 데이터베이스 연결이 설정되지 않았습니다.";
                    MessageBox.Show("데이터베이스를 먼저 상단에서 연결하고 불러와 주십시오.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string connStr = BuildConnectionString(false);
                SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder(connStr);
                sb.InitialCatalog = targetDb;

                this.Cursor = Cursors.WaitCursor;
                try
                {
                    using (SqlConnection conn = new SqlConnection(sb.ConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            if (lowerSql.StartsWith("select") || lowerSql.Contains("output") || lowerSql.Contains("returning"))
                            {
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dt);
                                }
                                _dgvQueryResult.DataSource = dt;
                                _lblQueryStatus.ForeColor = ColorEmerald;
                                _lblQueryStatus.Text = string.Format("조회 성공: {0}행 반환 완료.", dt.Rows.Count);
                            }
                            else
                            {
                                int affected = cmd.ExecuteNonQuery();
                                dt.Columns.Add("결과");
                                dt.Columns.Add("영향받은 행 수");
                                dt.Rows.Add("명령 실행 완료", affected.ToString() + "행");
                                _dgvQueryResult.DataSource = dt;

                                _lblQueryStatus.ForeColor = ColorEmerald;
                                _lblQueryStatus.Text = string.Format("명령 성공: {0}개의 행이 영향을 받았습니다.", affected);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _lblQueryStatus.ForeColor = ColorAlarm;
                    _lblQueryStatus.Text = "❌ 실행 에러: " + ex.Message;
                    MessageBox.Show("SQL 실행 중 오류가 발생했습니다:\n" + ex.Message, "쿼리 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }
    }

    public class PatientGroup
    {
        public string Name { get; set; }
        public string Jumin { get; set; }
        public int Count { get; set; }
        public string JuminEncrypt { get; set; }
        public string FamNm { get; set; }
    }

    public class RestoreSelectionForm : Form
    {
        public string SelectedName { get; private set; }
        public string SelectedJumin { get; private set; }
        public string SelectedEncrypt { get; private set; }
        public string SelectedFamNm { get; private set; }

        public RestoreSelectionForm(List<PatientGroup> candidates)
        {
            this.Text = "복구할 차트 주인 선택";
            this.Size = new Size(400, 210);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(15, 23, 42); // slate 900
            this.ForeColor = Color.FromArgb(248, 250, 252); // slate 50
            this.Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular);

            Label lbl = new Label 
            { 
                Text = "이 차트에 여러 환자의 처방이 섞여 있습니다.\n진짜 차트 주인을 선택해주세요:", 
                Location = new Point(20, 20), 
                Size = new Size(350, 45),
                ForeColor = Color.FromArgb(148, 163, 184) // slate 400
            };
            this.Controls.Add(lbl);

            ComboBox cmb = new ComboBox 
            { 
                Location = new Point(20, 75), 
                Size = new Size(340, 25), 
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(30, 41, 59), // slate 800
                ForeColor = Color.FromArgb(248, 250, 252) // slate 50
            };
            foreach (var c in candidates)
            {
                cmb.Items.Add(string.Format("{0} ({1} - 처방 {2}건)", c.Name, c.Jumin, c.Count));
            }
            cmb.SelectedIndex = 0;
            this.Controls.Add(cmb);

            Button btnOk = new Button 
            { 
                Text = "선택 정보로 복구", 
                Location = new Point(100, 120), 
                Size = new Size(180, 32), 
                FlatStyle = FlatStyle.Flat, 
                BackColor = Color.FromArgb(16, 185, 129), // Emerald 500
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Click += (s, e) => {
                var pg = candidates[cmb.SelectedIndex];
                SelectedName = pg.Name;
                SelectedJumin = pg.Jumin;
                SelectedEncrypt = pg.JuminEncrypt;
                SelectedFamNm = pg.FamNm;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            this.Controls.Add(btnOk);
        }
    }

    public class TroubleshooterForm : Form
    {
        private readonly MainForm _mainForm;
        private bool _isDemo;

        private static readonly Color ColorBgMain = Color.FromArgb(15, 23, 42);
        private static readonly Color ColorBgCard = Color.FromArgb(30, 41, 59);
        private static readonly Color ColorBorder = Color.FromArgb(51, 65, 85);
        private static readonly Color ColorTextMain = Color.FromArgb(248, 250, 252);
        private static readonly Color ColorTextSec = Color.FromArgb(148, 163, 184);
        private static readonly Color ColorIndigo = Color.FromArgb(99, 102, 241);
        private static readonly Color ColorEmerald = Color.FromArgb(16, 185, 129);
        private static readonly Color ColorAlarm = Color.FromArgb(239, 68, 68);

        // Controls
        private TabControl _tcScanner;
        private TabPage _tpDuplicate;
        private TabPage _tpMixed;
        private DataGridView _dgvDuplicate;
        private DataGridView _dgvMixed;
        private TextBox _txtChartNo;
        private Label _lblCurrentCustInfo;
        private DataGridView _dgvPrescDistribution;
        private ComboBox _cmbRestoreCandidates;
        private Button _btnRestoreMaster;
        private ComboBox _cmbMoveCandidates;
        private ComboBox _cmbDestCharts;
        private TextBox _txtCustomDestChart;
        private Button _btnMoveRx;
        private Button _btnQuickSolve;
        private Button _btnScanMixed;

        private class ComboItem
        {
            public string Text { get; set; }
            public object Value { get; set; }
            public override string ToString() { return Text; }
        }



        public TroubleshooterForm(MainForm mainForm, bool isDemo)
        {
            _mainForm = mainForm;
            _isDemo = isDemo;
            InitializeComponent();
        }

        public void ToggleDemoMode(bool isDemo)
        {
            _isDemo = isDemo;
            _txtChartNo.Text = "";
            _lblCurrentCustInfo.Text = "";
            if (_dgvPrescDistribution != null) _dgvPrescDistribution.DataSource = null;
            if (_dgvDuplicate != null) _dgvDuplicate.DataSource = null;
            if (_dgvMixed != null) _dgvMixed.DataSource = null;
            if (_cmbRestoreCandidates != null) _cmbRestoreCandidates.Items.Clear();
            if (_cmbMoveCandidates != null) _cmbMoveCandidates.Items.Clear();
            if (_cmbDestCharts != null) _cmbDestCharts.Items.Clear();
            if (_txtCustomDestChart != null) _txtCustomDestChart.Text = "";
            if (_tpDuplicate != null) _tpDuplicate.Text = "👤 개명 및 중복 차트";
            if (_tpMixed != null) _tpMixed.Text = "⚠️ 처방전 혼선/불일치";
        }

        private void InitializeComponent()
        {
            this.Text = "pm+helper - 차트 혼선 진단 및 처방전 분리 오류 해결 도구";
            this.Size = new Size(1100, 680);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = ColorBgMain;
            this.ForeColor = ColorTextMain;
            this.Font = new Font("맑은 고딕", 9.5F, FontStyle.Regular);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Split Container
            SplitContainer split = new SplitContainer();
            _mainForm._splitChartResolver = split; // 멤버 대입
            _mainForm._splitChartResolver.Dock = DockStyle.Fill;
            _mainForm._splitChartResolver.Orientation = Orientation.Vertical;
            _mainForm._splitChartResolver.Size = new Size(1050, 600);
            _mainForm._splitChartResolver.Panel1MinSize = 350;
            _mainForm._splitChartResolver.Panel2MinSize = 380;
            _mainForm._splitChartResolver.SplitterDistance = _mainForm._distChartResolver;
            _mainForm._splitChartResolver.BackColor = ColorBorder;
            this.Controls.Add(_mainForm._splitChartResolver);
            _mainForm._splitChartResolver.SplitterMoved += (s, e) => { _mainForm._distChartResolver = _mainForm._splitChartResolver.SplitterDistance; _mainForm.SaveConfig(); };

            // Left Panel (Scanner) -> Panel1에 배치
            Panel pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = ColorBgMain, Padding = new Padding(12) };
            _mainForm._splitChartResolver.Panel1.Controls.Add(pnlLeft);

            Label lblScannerTitle = new Label
            {
                Text = "🚨 차트 오류 유형별 전수 검사",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("맑은 고딕", 11F, FontStyle.Bold),
                ForeColor = ColorAlarm
            };

            _btnScanMixed = new Button
            {
                Text = "🔍 전체 테이블 검사 실행 (스캔)",
                Dock = DockStyle.Top,
                Height = 35,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            _btnScanMixed.FlatAppearance.BorderSize = 0;
            _btnScanMixed.Click += BtnScanMixed_Click;

            Panel pnlSpacer1 = new Panel { Dock = DockStyle.Top, Height = 10 };

            // TabControl for split views
            _tcScanner = new TabControl
            {
                Dock = DockStyle.Fill,
                BackColor = ColorBgMain,
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9F, FontStyle.Regular)
            };

            _tpDuplicate = new TabPage
            {
                Text = "👤 개명 및 중복 차트",
                BackColor = ColorBgCard
            };
            _tpMixed = new TabPage
            {
                Text = "⚠️ 처방전 혼선/불일치",
                BackColor = ColorBgCard
            };

            _tcScanner.TabPages.Add(_tpDuplicate);
            _tcScanner.TabPages.Add(_tpMixed);

            // Grids
            _dgvDuplicate = CreateScannerGrid();
            _dgvDuplicate.CellClick += DgvMixed_CellClick;

            Panel pnlDupBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 120,
                Padding = new Padding(0, 5, 0, 5),
                BackColor = ColorBgCard
            };
            Button btnExportExcel = new Button
            {
                Text = "📊 개명/중복 차트 목록 엑셀(CSV) 내보내기",
                Dock = DockStyle.Top,
                Height = 35,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorIndigo,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnExportExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.Click += BtnExportExcel_Click;
            pnlDupBottom.Controls.Add(btnExportExcel);

            Button btnDeleteEmptyDuplicates = new Button
            {
                Text = "🗑️ 처방 없는 중복 차트 일괄 삭제",
                Dock = DockStyle.Top,
                Height = 35,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnDeleteEmptyDuplicates.FlatAppearance.BorderSize = 0;
            btnDeleteEmptyDuplicates.Click += BtnDeleteEmptyDuplicates_Click;
            pnlDupBottom.Controls.Add(btnDeleteEmptyDuplicates);

            Button btnDeleteGhostCharts = new Button
            {
                Text = "🗑️ 이름 없는 유령 환자 차트 일괄 삭제",
                Dock = DockStyle.Top,
                Height = 35,
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold)
            };
            btnDeleteGhostCharts.FlatAppearance.BorderSize = 0;
            btnDeleteGhostCharts.Click += BtnDeleteGhostCharts_Click;
            pnlDupBottom.Controls.Add(btnDeleteGhostCharts);

            _tpDuplicate.Controls.Add(_dgvDuplicate);
            _tpDuplicate.Controls.Add(pnlDupBottom);
            pnlDupBottom.BringToFront();

            _dgvMixed = CreateScannerGrid();
            _dgvMixed.CellClick += DgvMixed_CellClick;
            _tpMixed.Controls.Add(_dgvMixed);


            pnlLeft.Controls.Add(_tcScanner);
            pnlLeft.Controls.Add(pnlSpacer1);
            pnlLeft.Controls.Add(_btnScanMixed);
            pnlLeft.Controls.Add(lblScannerTitle);

            // Right Panel (Console) -> Panel2에 배치
            Panel pnlRight = new Panel { Dock = DockStyle.Fill, BackColor = ColorBgMain, Padding = new Padding(12), AutoScroll = true };
            _mainForm._splitChartResolver.Panel2.Controls.Add(pnlRight);

            Label lblConsoleTitle = new Label
            {
                Text = "🛠️ 혼선 상세 분석 및 해결 조절판",
                Location = new Point(15, 12),
                Size = new Size(350, 25),
                Font = new Font("맑은 고딕", 11F, FontStyle.Bold),
                ForeColor = ColorEmerald
            };
            pnlRight.Controls.Add(lblConsoleTitle);

            // Group: Loaded Chart Info
            GroupBox gbChartInfo = new GroupBox
            {
                Text = "선택된 차트 정보 조회",
                Location = new Point(15, 45),
                Size = new Size(580, 100),
                ForeColor = ColorTextSec
            };
            pnlRight.Controls.Add(gbChartInfo);

            Label lblC1 = new Label { Text = "차트번호:", Location = new Point(15, 30), Size = new Size(70, 20), ForeColor = ColorTextSec };
            _txtChartNo = new TextBox { Location = new Point(90, 27), Size = new Size(120, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle };
            Button btnLoadChart = new Button { Text = "상세 조회", Location = new Point(220, 26), Size = new Size(80, 27), FlatStyle = FlatStyle.Flat, BackColor = ColorIndigo, ForeColor = Color.White };
            btnLoadChart.FlatAppearance.BorderSize = 0;
            btnLoadChart.Click += BtnLoadChart_Click;
            gbChartInfo.Controls.Add(lblC1);
            gbChartInfo.Controls.Add(_txtChartNo);
            gbChartInfo.Controls.Add(btnLoadChart);

            _lblCurrentCustInfo = new Label
            {
                Text = "조회된 차트 정보 없음 (조회해 주세요)",
                Location = new Point(15, 65),
                Size = new Size(550, 25),
                ForeColor = ColorTextMain,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            gbChartInfo.Controls.Add(_lblCurrentCustInfo);

            // Group: Prescription Distribution
            GroupBox gbDistribution = new GroupBox
            {
                Text = "차트 내부 처방전 분포 (혼선 현황)",
                Location = new Point(15, 155),
                Size = new Size(580, 150),
                ForeColor = ColorTextSec
            };
            pnlRight.Controls.Add(gbDistribution);

            _dgvPrescDistribution = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 28
            };
            _dgvPrescDistribution.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            _dgvPrescDistribution.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dgvPrescDistribution.DefaultCellStyle.BackColor = ColorBgCard;
            _dgvPrescDistribution.DefaultCellStyle.ForeColor = ColorTextMain;
            _dgvPrescDistribution.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            gbDistribution.Controls.Add(_dgvPrescDistribution);

            // Group: Step 1 (Restore Master)
            GroupBox gbStep1 = new GroupBox
            {
                Text = "1단계: 고객 마스터 정보 복구 (이름/주민번호 되돌리기)",
                Location = new Point(15, 315),
                Size = new Size(580, 95),
                ForeColor = ColorTextSec
            };
            pnlRight.Controls.Add(gbStep1);

            Label lblS1 = new Label { Text = "차트 주인 지정:", Location = new Point(15, 30), Size = new Size(110, 20), ForeColor = ColorTextSec };
            _cmbRestoreCandidates = new ComboBox { Location = new Point(130, 27), Size = new Size(220, 25), DropDownStyle = ComboBoxStyle.DropDownList, BackColor = ColorBgCard, ForeColor = ColorTextMain, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _btnRestoreMaster = new Button
            {
                Text = "선택 정보로 차트 주인 복구",
                Location = new Point(365, 25),
                Size = new Size(200, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorEmerald,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnRestoreMaster.FlatAppearance.BorderSize = 0;
            _btnRestoreMaster.Click += BtnRestoreMaster_Click;
            gbStep1.Controls.Add(lblS1);
            gbStep1.Controls.Add(_cmbRestoreCandidates);
            gbStep1.Controls.Add(_btnRestoreMaster);

            Label lblStep1Info = new Label
            {
                Text = "※ 선택된 환자의 실제 주민번호와 세대주명으로 고객 마스터를 원복합니다.",
                Location = new Point(15, 65),
                Size = new Size(550, 20),
                ForeColor = ColorTextSec,
                Font = new Font("맑은 고딕", 8.5F, FontStyle.Italic),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            gbStep1.Controls.Add(lblStep1Info);

            // Group: Step 2 (Move Prescriptions)
            GroupBox gbStep2 = new GroupBox
            {
                Text = "2단계: 처방전 분리 이동 (다른 환자의 처방전을 진짜 차트로 이관)",
                Location = new Point(15, 420),
                Size = new Size(580, 150),
                ForeColor = ColorTextSec
            };
            pnlRight.Controls.Add(gbStep2);

            Label lblS2 = new Label { Text = "이동할 처방:", Location = new Point(15, 30), Size = new Size(90, 20), ForeColor = ColorTextSec };
            _cmbMoveCandidates = new ComboBox { Location = new Point(110, 27), Size = new Size(240, 25), DropDownStyle = ComboBoxStyle.DropDownList, BackColor = ColorBgCard, ForeColor = ColorTextMain, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _cmbMoveCandidates.SelectedIndexChanged += CmbMoveCandidates_SelectedIndexChanged;
            gbStep2.Controls.Add(lblS2);
            gbStep2.Controls.Add(_cmbMoveCandidates);

            Label lblS3 = new Label { Text = "진짜 차트번호:", Location = new Point(15, 65), Size = new Size(95, 20), ForeColor = ColorTextSec };
            _cmbDestCharts = new ComboBox { Location = new Point(110, 62), Size = new Size(170, 25), DropDownStyle = ComboBoxStyle.DropDownList, BackColor = ColorBgCard, ForeColor = ColorTextMain, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _cmbDestCharts.SelectedIndexChanged += CmbDestCharts_SelectedIndexChanged;
            _txtCustomDestChart = new TextBox { Location = new Point(290, 62), Size = new Size(90, 25), BackColor = ColorBgMain, ForeColor = ColorTextMain, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            gbStep2.Controls.Add(lblS3);
            gbStep2.Controls.Add(_cmbDestCharts);
            gbStep2.Controls.Add(_txtCustomDestChart);

            Button btnFindTarget = new Button { Text = "차트 검색", Location = new Point(390, 61), Size = new Size(80, 27), FlatStyle = FlatStyle.Flat, BackColor = ColorIndigo, ForeColor = Color.White, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnFindTarget.FlatAppearance.BorderSize = 0;
            btnFindTarget.Click += BtnFindTarget_Click;
            gbStep2.Controls.Add(btnFindTarget);

            Button btnCreateNewChart = new Button { Text = "새로 만들기", Location = new Point(475, 61), Size = new Size(90, 27), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(245, 158, 11), ForeColor = Color.Black, Font = new Font("맑은 고딕", 9F, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnCreateNewChart.FlatAppearance.BorderSize = 0;
            btnCreateNewChart.Click += BtnCreateNewChart_Click;
            gbStep2.Controls.Add(btnCreateNewChart);

            _btnMoveRx = new Button
            {
                Text = "선택 처방전 및 수납 내역 이관 실행",
                Location = new Point(110, 105),
                Size = new Size(455, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = ColorAlarm,
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _btnMoveRx.FlatAppearance.BorderSize = 0;
            _btnMoveRx.Click += BtnMoveRx_Click;
            gbStep2.Controls.Add(_btnMoveRx);

            // Special: Quick Solve
            _btnQuickSolve = new Button
            {
                Text = "⚡ 천미선 복구 및 박복순 처방 분리 (원클릭 자동 해결 단축키)",
                Location = new Point(15, 580),
                Size = new Size(580, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(79, 70, 229), // Indigo 600
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold),
                Visible = false
            };
            _btnQuickSolve.FlatAppearance.BorderSize = 0;
            _btnQuickSolve.Click += BtnQuickSolve_Click;
            pnlRight.Controls.Add(_btnQuickSolve);

            // Dynamically resize widths of right panel GroupBoxes & QuickSolve button
            // This prevents AutoScroll conflict with AnchorStyles.Right (WinForms layout bug).
            pnlRight.Resize += (s, e) =>
            {
                int padding = 15;
                int targetWidth = pnlRight.ClientSize.Width - (padding * 2);
                if (targetWidth < 200) targetWidth = 200;

                gbChartInfo.Width = targetWidth;
                gbDistribution.Width = targetWidth;
                gbStep1.Width = targetWidth;
                gbStep2.Width = targetWidth;
                _btnQuickSolve.Width = targetWidth;
            };
            pnlRight.PerformLayout();
            int initialRightWidth = pnlRight.ClientSize.Width - 30;
            if (initialRightWidth < 200) initialRightWidth = 200;
            gbChartInfo.Width = initialRightWidth;
            gbDistribution.Width = initialRightWidth;
            gbStep1.Width = initialRightWidth;
            gbStep2.Width = initialRightWidth;
            _btnQuickSolve.Width = initialRightWidth;
        }

        public void LoadScannerGrid()
        {
            if (_isDemo)
            {
                ScanMixedDemo();
            }
            else
            {
                ScanMixedProduction();
            }
        }

        private void BtnScanMixed_Click(object sender, EventArgs e)
        {
            if (_btnScanMixed != null)
            {
                _btnScanMixed.Enabled = false;
                _btnScanMixed.Text = "⏳ 스캔 진행 중 (데이터 분석 중)...";
            }
            LoadScannerGrid();
        }

        private void ScanMixedDemo()
        {
            // Find duplicate active Jumin encrypts in mock customer master
            var activeJumins = new Dictionary<string, List<string>>();
            foreach (var cust in _mainForm._mockCustList)
            {
                if (cust.CusAct == "1" && !string.IsNullOrEmpty(cust.PatJuminNo) && char.IsDigit(cust.PatJuminNo[0]))
                {
                    if (!activeJumins.ContainsKey(cust.JuminEncrypt))
                    {
                        activeJumins[cust.JuminEncrypt] = new List<string>();
                    }
                    if (!activeJumins[cust.JuminEncrypt].Contains(cust.ChrtNo))
                    {
                        activeJumins[cust.JuminEncrypt].Add(cust.ChrtNo);
                    }
                }
            }

            var duplicateJumins = new HashSet<string>();
            foreach (var kvp in activeJumins)
            {
                if (kvp.Value.Count > 1)
                {
                    duplicateJumins.Add(kvp.Key);
                }
            }

            DataTable dtDup = new DataTable();
            dtDup.Columns.Add("차트번호");
            dtDup.Columns.Add("환자명");
            dtDup.Columns.Add("주민번호");
            dtDup.Columns.Add("총 처방수", typeof(int));
            dtDup.Columns.Add("세대주");

            foreach (var cust in _mainForm._mockCustList)
            {
                if (duplicateJumins.Contains(cust.JuminEncrypt) && cust.CusAct == "1")
                {
                    int rxCount = _mainForm._mockRxList.FindAll(rx => rx.ChrtNo == cust.ChrtNo).Count;
                    dtDup.Rows.Add(cust.ChrtNo, cust.PatNm, cust.PatJuminNo, rxCount, cust.FamNm);
                }
            }

            DataTable dtMixed = new DataTable();
            dtMixed.Columns.Add("차트번호");
            dtMixed.Columns.Add("환자수", typeof(int));
            dtMixed.Columns.Add("주민수", typeof(int));
            dtMixed.Columns.Add("환자1");
            dtMixed.Columns.Add("환자2");
            dtMixed.Columns.Add("총 처방수", typeof(int));

            var grouped = new Dictionary<string, List<MainForm.MockRx>>();
            foreach (var rx in _mainForm._mockRxList)
            {
                string key = string.IsNullOrEmpty(rx.ChrtNo) ? "" : rx.ChrtNo.Trim();
                if (!grouped.ContainsKey(key)) grouped[key] = new List<MainForm.MockRx>();
                grouped[key].Add(rx);
            }

            foreach (var kvp in grouped)
            {
                List<string> names = new List<string>();
                List<string> jumins = new List<string>();
                foreach (var rx in kvp.Value)
                {
                    if (!names.Contains(rx.PatNm)) names.Add(rx.PatNm);
                    if (!jumins.Contains(rx.PatJuminNo)) jumins.Add(rx.PatJuminNo);
                }

                bool isMismatched = false;
                var cust = _mainForm._mockCustList.Find(c => c.ChrtNo == kvp.Key && c.CusAct == "1");
                if (cust != null)
                {
                    string cleanRxJumin = kvp.Value[0].PatJuminNo.Replace("-", "");
                    string cleanCustJumin = cust.PatJuminNo.Replace("-", "");
                    if (cleanRxJumin.Length >= 7 && cleanCustJumin.Length >= 7 && cleanRxJumin.Substring(0, 7) != cleanCustJumin.Substring(0, 7))
                    {
                        isMismatched = true;
                    }
                }

                if (jumins.Count > 1 || isMismatched)
                {
                    dtMixed.Rows.Add(
                        kvp.Key,
                        names.Count,
                        jumins.Count,
                        names.Count > 0 ? names[0] : "",
                        names.Count > 1 ? names[1] : (cust != null ? cust.PatNm : ""),
                        kvp.Value.Count
                    );
                }
            }

            _dgvDuplicate.DataSource = dtDup;
            _dgvMixed.DataSource = dtMixed;
            AdjustScannerGridWidths();
            _tpDuplicate.Text = string.Format("👤 개명 및 중복 차트 ({0}건)", dtDup.Rows.Count);
            _tpMixed.Text = string.Format("⚠️ 처방전 혼선/불일치 ({0}건)", dtMixed.Rows.Count);

            if (_btnScanMixed != null)
            {
                _btnScanMixed.Enabled = true;
                _btnScanMixed.Text = "🔍 전체 테이블 검사 실행 (스캔)";
            }
            MessageBox.Show("[데모] 차트 전수 검사(스캔)가 성공적으로 완료되었습니다.", "스캔 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ScanMixedProduction()
        {
            string connStr = _mainForm.BuildConnectionString(false);
            this.Cursor = Cursors.WaitCursor;

            // Run database queries in background thread
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    DataTable dtDup = new DataTable();
                    DataTable dtMixed = new DataTable();

                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();

                        // Query 1: Duplicate / Name Change charts
                        string sqlDup = @"
                            WITH DuplicateJumins AS (
                                SELECT jumin_encrypt
                                FROM tbsit000_01
                                WHERE cusact = '1' 
                                  AND jumin_no LIKE '[0-9]%'
                                GROUP BY jumin_encrypt
                                HAVING COUNT(DISTINCT chrtno) > 1
                            )
                            SELECT m.chrtno as [차트번호],
                                   m.pat_nm as [환자명],
                                   m.jumin_no as [주민번호],
                                   (SELECT COUNT(*) FROM tbsid040_03 r WHERE r.chrtno = m.chrtno) as [총 처방수],
                                   m.fam_nm as [세대주]
                            FROM tbsit000_01 m
                            WHERE m.jumin_encrypt IN (SELECT jumin_encrypt FROM DuplicateJumins)
                              AND m.cusact = '1'
                            ORDER BY [주민번호] ASC, [총 처방수] DESC";

                        using (SqlCommand cmd = new SqlCommand(sqlDup, conn))
                        {
                            cmd.CommandTimeout = 300;
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dtDup);
                            }
                        }

                        // Query 2: Mixed / Mismatched Rx charts
                        string sqlMixed = @"
                            WITH TargetCharts AS (
                                SELECT COALESCE(NULLIF(LTRIM(RTRIM(r.chrtno)), ''), '') as chrtno, 
                                       COUNT(DISTINCT r.pat_nm) as rx_names, 
                                       COUNT(DISTINCT r.pat_jumin_no) as rx_jumins,
                                       COUNT(*) as rx_count,
                                       MIN(r.pat_nm) as rx_name_min,
                                       MAX(r.pat_nm) as rx_name_max
                                FROM tbsid040_03 r
                                GROUP BY COALESCE(NULLIF(LTRIM(RTRIM(r.chrtno)), ''), '')
                            ),
                            MismatchedCharts AS (
                                SELECT DISTINCT COALESCE(NULLIF(LTRIM(RTRIM(r.chrtno)), ''), '') as chrtno, m.pat_nm as master_name
                                FROM tbsid040_03 r
                                JOIN tbsit000_01 m ON COALESCE(NULLIF(LTRIM(RTRIM(r.chrtno)), ''), '') = m.chrtno
                                WHERE m.cusact = '1'
                                  AND (SUBSTRING(REPLACE(r.pat_jumin_no, '-', ''), 1, 7) <> SUBSTRING(REPLACE(m.jumin_no, '-', ''), 1, 7))
                            )
                            SELECT tc.chrtno as [차트번호],
                                   tc.rx_names as [환자수],
                                   tc.rx_jumins as [주민수],
                                   tc.rx_name_min as [환자1],
                                   COALESCE(
                                       CASE WHEN tc.rx_names > 1 THEN (SELECT MAX(r2.pat_nm) FROM tbsid040_03 r2 WHERE COALESCE(NULLIF(LTRIM(RTRIM(r2.chrtno)), ''), '') = tc.chrtno AND r2.pat_nm <> tc.rx_name_min) ELSE NULL END,
                                       mc.master_name,
                                       (SELECT TOP 1 m2.pat_nm FROM tbsit000_01 m2 WHERE m2.jumin_encrypt = (SELECT TOP 1 j2.jumin_encrypt FROM tbsit000_01 j2 WHERE COALESCE(NULLIF(LTRIM(RTRIM(j2.chrtno)), ''), '') = tc.chrtno) AND m2.chrtno <> tc.chrtno AND m2.cusact = '1'),
                                       tc.rx_name_max
                                   ) as [환자2],
                                   tc.rx_count as [총 처방수]
                            FROM TargetCharts tc
                            LEFT JOIN MismatchedCharts mc ON tc.chrtno = mc.chrtno
                            WHERE tc.rx_jumins > 1 
                               OR mc.chrtno IS NOT NULL
                            ORDER BY [총 처방수] DESC";

                        using (SqlCommand cmd = new SqlCommand(sqlMixed, conn))
                        {
                            cmd.CommandTimeout = 300;
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dtMixed);
                            }
                        }
                    }

                    // Update UI on main thread
                    this.BeginInvoke((Action)(() =>
                    {
                        _dgvDuplicate.DataSource = dtDup;
                        _dgvMixed.DataSource = dtMixed;
                        AdjustScannerGridWidths();
                        _tpDuplicate.Text = string.Format("👤 개명 및 중복 차트 ({0}건)", dtDup.Rows.Count);
                        _tpMixed.Text = string.Format("⚠️ 처방전 혼선/불일치 ({0}건)", dtMixed.Rows.Count);
                        this.Cursor = Cursors.Default;
                        if (_btnScanMixed != null)
                        {
                            _btnScanMixed.Enabled = true;
                            _btnScanMixed.Text = "🔍 전체 테이블 검사 실행 (스캔)";
                        }
                        MessageBox.Show("차트 전수 검사(스캔)가 성공적으로 완료되었습니다.", "스캔 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        this.Cursor = Cursors.Default;
                        if (_btnScanMixed != null)
                        {
                            _btnScanMixed.Enabled = true;
                            _btnScanMixed.Text = "🔍 전체 테이블 검사 실행 (스캔)";
                        }
                        MessageBox.Show("혼선 차트 스캔 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private DataGridView CreateScannerGrid()
        {
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = ColorBgCard,
                ForeColor = ColorTextMain,
                GridColor = ColorBorder,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = ColorBgMain;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold);
            dgv.DefaultCellStyle.BackColor = ColorBgCard;
            dgv.DefaultCellStyle.ForeColor = ColorTextMain;
            dgv.DefaultCellStyle.SelectionBackColor = ColorIndigo;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            return dgv;
        }

        private void DgvMixed_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridView dgv = (DataGridView)sender;
            var val = dgv.Rows[e.RowIndex].Cells[0].Value;
            if (val != null)
            {
                _txtChartNo.Text = val.ToString().Trim();
                BtnLoadChart_Click(null, null);
            }
        }

        private void AdjustScannerGridWidths()
        {
            if (_dgvDuplicate.Columns.Count > 0)
            {
                _dgvDuplicate.Columns[0].Width = 85;  // 차트번호
                _dgvDuplicate.Columns[1].Width = 75;  // 환자명
                _dgvDuplicate.Columns[2].Width = 100; // 주민번호
                _dgvDuplicate.Columns[3].Width = 70;  // 총 처방수
                _dgvDuplicate.Columns[4].Width = 70;  // 세대주

                ApplyContentSizedColumns(_dgvDuplicate);

                if (_dgvDuplicate.Columns.Count > 2)
                {
                    _dgvDuplicate.Sort(_dgvDuplicate.Columns[2], System.ComponentModel.ListSortDirection.Ascending);
                    ApplyContentSizedColumns(_dgvDuplicate);
                }
            }
            if (_dgvMixed.Columns.Count > 0)
            {
                _dgvMixed.Columns[0].Width = 85;  // 차트번호
                _dgvMixed.Columns[1].Width = 50;  // 환자수
                _dgvMixed.Columns[2].Width = 50;  // 주민수
                _dgvMixed.Columns[3].Width = 75;  // 환자1
                _dgvMixed.Columns[4].Width = 75;  // 환자2
                _dgvMixed.Columns[5].Width = 70;  // 총 처방수
                ApplyContentSizedColumns(_dgvMixed);
            }
            AdjustChartResolverSplitFromGrids();
        }

        private void ApplyContentSizedColumns(DataGridView dgv)
        {
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Visible) continue;

                int targetWidth = TextRenderer.MeasureText(col.HeaderText, dgv.ColumnHeadersDefaultCellStyle.Font).Width + 32;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    object value = row.Cells[col.Index].FormattedValue;
                    string text = value == null ? "" : value.ToString();
                    int cellWidth = TextRenderer.MeasureText(text, dgv.DefaultCellStyle.Font).Width + 28;
                    if (cellWidth > targetWidth) targetWidth = cellWidth;
                }

                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                col.MinimumWidth = Math.Max(60, targetWidth);
                col.Width = col.MinimumWidth;
            }
            dgv.Invalidate();
        }

        private void AdjustChartResolverSplitFromGrids()
        {
            if (_mainForm._splitChartResolver == null) return;

            int gridWidth = GetGridContentWidth(_dgvDuplicate);
            gridWidth = Math.Max(gridWidth, GetGridContentWidth(_dgvMixed));

            int leftChrome = 48;
            int desiredLeft = Math.Max(420, gridWidth + leftChrome);
            int totalWidth = _mainForm._splitChartResolver.ClientSize.Width;
            int rightMin = Math.Max(380, _mainForm._splitChartResolver.Panel2MinSize);
            int maxLeft = Math.Max(_mainForm._splitChartResolver.Panel1MinSize, totalWidth - rightMin);
            int newDistance = Math.Min(desiredLeft, maxLeft);

            if (newDistance >= _mainForm._splitChartResolver.Panel1MinSize)
            {
                try
                {
                    _mainForm._splitChartResolver.SplitterDistance = newDistance;
                    _mainForm._distChartResolver = newDistance;
                }
                catch { }
            }
        }

        private int GetGridContentWidth(DataGridView dgv)
        {
            if (dgv == null || dgv.Columns.Count == 0) return 0;

            int width = 0;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Visible)
                {
                    width += col.Width;
                }
            }
            return width;
        }

        private void BtnLoadChart_Click(object sender, EventArgs e)
        {
            string chrtno = _txtChartNo.Text.Trim();

            if (_isDemo)
            {
                LoadChartDemo(chrtno);
            }
            else
            {
                LoadChartProduction(chrtno);
            }
        }

        private void LoadChartDemo(string chrtno)
        {
            // 1. Customer Info
            var cust = _mainForm._mockCustList.Find(c => c.ChrtNo == chrtno);
            if (cust != null)
            {
                _lblCurrentCustInfo.Text = string.Format("현재 마스터: {0} ({1}) / 세대주: {2} / 상태: {3}", cust.PatNm, cust.PatJuminNo, cust.FamNm, cust.CusAct == "1" ? "활성" : "중지");
            }
            else
            {
                _lblCurrentCustInfo.Text = "고객 마스터에 존재하지 않는 임의의 차트입니다.";
            }

            // 2. Prescriptions Distribution
            DataTable dt = new DataTable();
            dt.Columns.Add("환자 이름");
            dt.Columns.Add("주민번호");
            dt.Columns.Add("처방 건수", typeof(int));

            var dict = new Dictionary<string, PatientGroup>();
            foreach (var rx in _mainForm._mockRxList)
            {
                if (rx.ChrtNo == chrtno)
                {
                    string key = rx.PatNm + "|" + rx.PatJuminNo;
                    if (!dict.ContainsKey(key))
                    {
                        dict[key] = new PatientGroup { Name = rx.PatNm, Jumin = rx.PatJuminNo, Count = 0, JuminEncrypt = rx.JuminEncrypt, FamNm = rx.PatNm == "천미선" ? "백승현" : "임광묵" };
                    }
                    dict[key].Count++;
                }
            }

            _cmbRestoreCandidates.Items.Clear();
            _cmbMoveCandidates.Items.Clear();

            foreach (var kvp in dict.Values)
            {
                dt.Rows.Add(kvp.Name, kvp.Jumin, kvp.Count);

                ComboItem item = new ComboItem
                {
                    Text = string.Format("{0} ({1} - 처방 {2}건)", kvp.Name, kvp.Jumin, kvp.Count),
                    Value = kvp
                };
                _cmbRestoreCandidates.Items.Add(item);
                _cmbMoveCandidates.Items.Add(item);
            }

            _dgvPrescDistribution.DataSource = dt;
            if (_dgvPrescDistribution.Columns.Count > 0)
            {
                _dgvPrescDistribution.Columns[0].Width = 120;
                _dgvPrescDistribution.Columns[1].Width = 180;
                _dgvPrescDistribution.Columns[2].Width = 100;
            }

            if (_cmbRestoreCandidates.Items.Count > 0) _cmbRestoreCandidates.SelectedIndex = 0;
            if (_cmbMoveCandidates.Items.Count > 0) _cmbMoveCandidates.SelectedIndex = 0;

            // Show Quick Solve button if Case 1
            _btnQuickSolve.Visible = (chrtno == "0000184791");
        }

        private void LoadChartProduction(string chrtno)
        {
            string connStr = _mainForm.BuildConnectionString(false);
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();

                    // 1. Customer master
                    string queryCust = "SELECT pat_nm, jumin_no, fam_nm, cusact FROM tbsit000_01 WHERE chrtno = @chrtno";
                    using (SqlCommand cmd = new SqlCommand(queryCust, conn))
                    {
                        cmd.Parameters.AddWithValue("@chrtno", chrtno);
                        using (SqlDataReader r = cmd.ExecuteReader())
                        {
                            if (r.Read())
                            {
                                _lblCurrentCustInfo.Text = string.Format("현재 마스터: {0} ({1}) / 세대주: {2} / 상태: {3}",
                                    r["pat_nm"], r["jumin_no"], r["fam_nm"], r["cusact"].ToString() == "1" ? "활성" : "중지");
                            }
                            else
                            {
                                _lblCurrentCustInfo.Text = "고객 마스터에 존재하지 않는 임의의 차트입니다.";
                            }
                        }
                    }

                    // 2. Prescriptions Distribution
                    string queryRx = @"
                        SELECT pat_nm, pat_jumin_no, COUNT(*) as rx_count, MIN(jumin_encrypt) as jumin_encrypt, MIN(fam_nm) as fam_nm
                        FROM tbsid040_03
                        WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @chrtno
                        GROUP BY pat_nm, pat_jumin_no";

                    DataTable dt = new DataTable();
                    dt.Columns.Add("환자 이름");
                    dt.Columns.Add("주민번호");
                    dt.Columns.Add("처방 건수", typeof(int));

                    _cmbRestoreCandidates.Items.Clear();
                    _cmbMoveCandidates.Items.Clear();

                    using (SqlCommand cmd = new SqlCommand(queryRx, conn))
                    {
                        cmd.Parameters.AddWithValue("@chrtno", chrtno);
                        using (SqlDataReader r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                string name = r["pat_nm"].ToString();
                                string jumin = r["pat_jumin_no"].ToString();
                                int count = Convert.ToInt32(r["rx_count"]);
                                string encrypt = r["jumin_encrypt"].ToString();
                                string famNm = r["fam_nm"].ToString();

                                dt.Rows.Add(name, jumin, count);

                                PatientGroup pg = new PatientGroup { Name = name, Jumin = jumin, Count = count, JuminEncrypt = encrypt, FamNm = famNm };
                                ComboItem item = new ComboItem
                                {
                                    Text = string.Format("{0} ({1} - 처방 {2}건)", name, jumin, count),
                                    Value = pg
                                };
                                _cmbRestoreCandidates.Items.Add(item);
                                _cmbMoveCandidates.Items.Add(item);
                            }
                        }
                    }

                    _dgvPrescDistribution.DataSource = dt;
                    if (_dgvPrescDistribution.Columns.Count > 0)
                    {
                        _dgvPrescDistribution.Columns[0].Width = 120;
                        _dgvPrescDistribution.Columns[1].Width = 180;
                        _dgvPrescDistribution.Columns[2].Width = 100;
                    }

                    if (_cmbRestoreCandidates.Items.Count > 0) _cmbRestoreCandidates.SelectedIndex = 0;
                    if (_cmbMoveCandidates.Items.Count > 0) _cmbMoveCandidates.SelectedIndex = 0;

                    _btnQuickSolve.Visible = (chrtno == "0000184791");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("차트 정보 조회 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRestoreMaster_Click(object sender, EventArgs e)
        {
            if (_cmbRestoreCandidates.SelectedItem == null) return;
            ComboItem selected = (ComboItem)_cmbRestoreCandidates.SelectedItem;
            PatientGroup pg = (PatientGroup)selected.Value;

            string chrtno = _txtChartNo.Text.Trim();
            DialogResult dr = MessageBox.Show(
                string.Format("차트번호 [{0}]의 고객 마스터 정보를 [{1}] 님의 정보로 복구하시겠습니까?\n\n- 복구할 이름: {1}\n- 복구할 주민번호: {2}\n- 세대주: {3}", 
                    chrtno, pg.Name, pg.Jumin, pg.FamNm),
                "고객 정보 복구 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (dr != DialogResult.Yes) return;

            _mainForm.RestoreMaster(chrtno, pg);
            LoadChartDetails(chrtno);
        }





        private void CmbMoveCandidates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbMoveCandidates.SelectedItem == null) return;
            ComboItem selected = (ComboItem)_cmbMoveCandidates.SelectedItem;
            PatientGroup pg = (PatientGroup)selected.Value;

            // Clear destination charts
            _cmbDestCharts.Items.Clear();
            _txtCustomDestChart.Text = "";

            if (_isDemo)
            {
                var matches = _mainForm._mockCustList.FindAll(c => 
                    (c.PatNm == pg.Name || c.JuminEncrypt == pg.JuminEncrypt) && 
                    c.ChrtNo != _txtChartNo.Text.Trim() && 
                    c.CusAct == "1"
                );
                foreach (var m in matches)
                {
                    _cmbDestCharts.Items.Add(new ComboItem { Text = string.Format("{0} ({1}, {2})", m.ChrtNo, m.PatNm, m.PatJuminNo), Value = m.ChrtNo });
                }
            }
            else
            {
                string connStr = _mainForm.BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT chrtno, pat_nm, jumin_no FROM tbsit000_01 WHERE (pat_nm = @pat_nm OR jumin_encrypt = @jumin_encrypt) AND chrtno <> @current_chrtno AND cusact = '1'";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@pat_nm", pg.Name);
                            cmd.Parameters.AddWithValue("@jumin_encrypt", pg.JuminEncrypt);
                            cmd.Parameters.AddWithValue("@current_chrtno", _txtChartNo.Text.Trim());
                            using (SqlDataReader r = cmd.ExecuteReader())
                            {
                                while (r.Read())
                                {
                                    string chrt = r["chrtno"].ToString();
                                    string name = r["pat_nm"].ToString();
                                    string jumin = r["jumin_no"].ToString();
                                    _cmbDestCharts.Items.Add(new ComboItem { Text = string.Format("{0} ({1}, {2})", chrt, name, jumin), Value = chrt });
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Fail silently, user can input manually
                }
            }

            if (_cmbDestCharts.Items.Count > 0)
            {
                _cmbDestCharts.SelectedIndex = 0;
            }
        }

        private void CmbDestCharts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_cmbDestCharts.SelectedItem != null)
            {
                ComboItem selected = (ComboItem)_cmbDestCharts.SelectedItem;
                _txtCustomDestChart.Text = selected.Value.ToString();
            }
        }

        private void BtnFindTarget_Click(object sender, EventArgs e)
        {
            if (_cmbMoveCandidates.SelectedItem == null) return;
            ComboItem selected = (ComboItem)_cmbMoveCandidates.SelectedItem;
            PatientGroup pg = (PatientGroup)selected.Value;

            CmbMoveCandidates_SelectedIndexChanged(null, null);
            MessageBox.Show(string.Format("[{0}] 환자명과 일치하는 다른 고객 차트 목록을 조회하여 리스트에 채웠습니다.", pg.Name), "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnMoveRx_Click(object sender, EventArgs e)
        {
            if (_cmbMoveCandidates.SelectedItem == null) return;
            ComboItem selected = (ComboItem)_cmbMoveCandidates.SelectedItem;
            PatientGroup pg = (PatientGroup)selected.Value;

            string destChart = _txtCustomDestChart.Text.Trim();
            if (string.IsNullOrEmpty(destChart))
            {
                MessageBox.Show("처방전을 이동시킬 진짜 차트번호를 입력하거나 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string srcChart = _txtChartNo.Text.Trim();
            if (srcChart == destChart)
            {
                MessageBox.Show("출발지 차트번호와 목적지 차트번호가 동일합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(
                string.Format("차트 [{0}]의 [{1}] 님 처방 데이터 {2}건을 [{3}] 차트로 이관하시겠습니까?\n(이 작업은 해당 처방과 관련된 수납/매출 원장 내역도 함께 이관합니다.)",
                    srcChart, pg.Name, pg.Count, destChart),
                "처방전 및 수납 내역 이관 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            if (_isDemo)
            {
                MoveRxDemo(srcChart, pg.Name, pg.Jumin, destChart);
            }
            else
            {
                MoveRxProduction(srcChart, pg.Name, pg.Jumin, destChart);
            }
        }

        private void MoveRxDemo(string srcChart, string patNm, string jumin, string destChart)
        {
            bool destExists = _mainForm._mockCustList.Exists(c => c.ChrtNo == destChart);
            string destName = patNm; // default to current name

            if (!destExists)
            {
                DialogResult drCheck = MessageBox.Show(
                    string.Format("이송 목적지 차트번호 [{0}]은(는) 고객 마스터(데모)에 존재하지 않습니다.\n\n" +
                                  "해당 차트번호로 '{1}' 님(주민번호: {2})의 신규 고객 정보를 등록(마스터 생성)하고 처방전을 이관하시겠습니까?", 
                                  destChart, patNm, _mainForm.FormatJuminFull(jumin)),
                    "신규 고객 마스터 등록 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (drCheck == DialogResult.Yes)
                {
                    // Find lookup details from mock rx list
                    string encrypt = "ENC_MOCK_" + patNm;
                    string famNm = (patNm == "천미선") ? "백승현" : (patNm == "김현숙") ? "윤창식" : "임광묵";
                    var lookup = _mainForm._mockRxList.Find(rx => rx.ChrtNo == srcChart && rx.PatNm == patNm && rx.PatJuminNo == jumin);
                    if (lookup != null)
                    {
                        encrypt = lookup.JuminEncrypt;
                    }

                    // Create mock customer master
                    var newCust = new MainForm.MockCust
                    {
                        ChrtNo = destChart,
                        PatNm = patNm,
                        PatJuminNo = jumin,
                        JuminNo = jumin.Replace("-", ""),
                        JuminEncrypt = encrypt,
                        FamNm = famNm,
                        Phone = "010-1234-5678",
                        Address = "서울시 강남구 역삼동",
                        FirstVisit = DateTime.Now.ToString("yyyy-MM-dd"),
                        CusAct = "1"
                    };
                    _mainForm._mockCustList.Add(newCust);
                }
                else
                {
                    return; // Cancel transfer
                }
            }
            else
            {
                // Retrieve destination name from mock customer
                var destCust = _mainForm._mockCustList.Find(c => c.ChrtNo == destChart);
                if (destCust != null)
                {
                    destName = destCust.PatNm;
                }
            }

            int rxCount = 0;
            foreach (var rx in _mainForm._mockRxList)
            {
                if (rx.ChrtNo == srcChart && rx.PatNm == patNm && rx.PatJuminNo == jumin)
                {
                    rx.ChrtNo = destChart;
                    rx.PatNm = destName; // align name
                    rxCount++;
                }
            }

            // Deactivate source chart in mock if it shares same Jumin and has no rx left
            var srcCust = _mainForm._mockCustList.Find(c => c.ChrtNo == srcChart);
            var destCustCheck = _mainForm._mockCustList.Find(c => c.ChrtNo == destChart);
            if (srcCust != null && destCustCheck != null && srcCust.JuminEncrypt == destCustCheck.JuminEncrypt)
            {
                bool hasRxLeft = _mainForm._mockRxList.Exists(rx => rx.ChrtNo == srcChart);
                if (!hasRxLeft)
                {
                    srcCust.CusAct = "0"; // deactivate
                }
            }

            MessageBox.Show(string.Format("[데모] {0}건의 처방 및 수납 정보가 차트 [{1}]({2})로 안전하게 이관되었습니다.", rxCount, destChart, destName), "성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadChartDetails(srcChart);
        }

        private void MoveRxProduction(string srcChart, string patNm, string jumin, string destChart)
        {
            string connStr = _mainForm.BuildConnectionString(false);

            // Check if destination chart exists in customer master
            bool destExists = false;
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string checkSql = "SELECT COUNT(*) FROM tbsit000_01 WHERE chrtno = @destChart";
                    using (SqlCommand checkCmd = new SqlCommand(checkSql, conn))
                    {
                        checkCmd.Parameters.AddWithValue("@destChart", destChart);
                        int count = Convert.ToInt32(checkCmd.ExecuteScalar());
                        destExists = (count > 0);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("목적지 차트번호 존재 여부 확인 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            bool createNewMaster = false;
            if (!destExists)
            {
                DialogResult drCheck = MessageBox.Show(
                    string.Format("이송 목적지 차트번호 [{0}]은(는) 고객 마스터(tbsit000_01)에 존재하지 않습니다.\n\n" +
                                  "해당 차트번호로 '{1}' 님(주민번호: {2})의 신규 고객 정보를 등록(마스터 생성)하고 처방전을 이관하시겠습니까?", 
                                  destChart, patNm, _mainForm.FormatJuminFull(jumin)),
                    "신규 고객 마스터 등록 확인",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                if (drCheck == DialogResult.Yes)
                {
                    createNewMaster = true;
                }
                else
                {
                    return; // Cancel transfer
                }
            }

            this.Cursor = Cursors.WaitCursor;

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                using (SqlTransaction trans = conn.BeginTransaction())
                {
                    try
                    {
                        string destName = patNm; // default to current name

                        // If new master creation is requested
                        if (createNewMaster)
                        {
                            // Retrieve encrypt and fam_nm from tbsid040_03
                            string encrypt = "";
                            string famNm = "";
                            string lookupSql = @"
                                SELECT TOP 1 jumin_encrypt, fam_nm 
                                FROM tbsid040_03 
                                WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @srcChart AND pat_nm = @pat_nm AND pat_jumin_no = @pat_jumin_no";
                            using (SqlCommand lookupCmd = new SqlCommand(lookupSql, conn, trans))
                            {
                                lookupCmd.Parameters.AddWithValue("@srcChart", srcChart);
                                lookupCmd.Parameters.AddWithValue("@pat_nm", patNm);
                                lookupCmd.Parameters.AddWithValue("@pat_jumin_no", jumin);
                                using (SqlDataReader r = lookupCmd.ExecuteReader())
                                {
                                    if (r.Read())
                                    {
                                        encrypt = r["jumin_encrypt"].ToString();
                                        famNm = r["fam_nm"].ToString();
                                    }
                                }
                            }

                            // If not found in rx, look up in TEMP_MAPPING_CHRTNO
                            if (string.IsNullOrEmpty(encrypt) && _mainForm.TableExists(conn, "TEMP_MAPPING_CHRTNO", trans))
                            {
                                string tempSql = "SELECT JUMIN_ENCRYPT FROM TEMP_MAPPING_CHRTNO WHERE pat_nm = @pat_nm AND chrtno = @srcChart";
                                using (SqlCommand tempCmd = new SqlCommand(tempSql, conn, trans))
                                {
                                    tempCmd.Parameters.AddWithValue("@pat_nm", patNm);
                                    tempCmd.Parameters.AddWithValue("@srcChart", srcChart);
                                    var res = tempCmd.ExecuteScalar();
                                    if (res != null) encrypt = res.ToString();
                                }
                            }

                            // Insert new customer master
                            string insertSql = @"
                                INSERT INTO tbsit000_01 (chrtno, pat_seq, pat_nm, jumin_no, jumin_encrypt, fam_nm, cusact, proc_dtime)
                                VALUES (@chrtno, 1, @pat_nm, @jumin_no, @jumin_encrypt, @fam_nm, '1', @proc_dtime)";
                            using (SqlCommand insCmd = new SqlCommand(insertSql, conn, trans))
                            {
                                insCmd.Parameters.AddWithValue("@chrtno", destChart);
                                insCmd.Parameters.AddWithValue("@pat_nm", patNm);
                                insCmd.Parameters.AddWithValue("@jumin_no", jumin);
                                insCmd.Parameters.AddWithValue("@jumin_encrypt", encrypt);
                                insCmd.Parameters.AddWithValue("@fam_nm", famNm);
                                insCmd.Parameters.AddWithValue("@proc_dtime", DateTime.Now.ToString("yyyyMMddHHmmss"));
                                insCmd.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            // Retrieve destination master name
                            string masterNameSql = "SELECT pat_nm FROM tbsit000_01 WHERE chrtno = @destChart";
                            using (SqlCommand mnCmd = new SqlCommand(masterNameSql, conn, trans))
                            {
                                mnCmd.Parameters.AddWithValue("@destChart", destChart);
                                var res = mnCmd.ExecuteScalar();
                                if (res != null && res != DBNull.Value)
                                {
                                    destName = res.ToString();
                                }
                            }
                        }

                        // 1. Get all drug_seq values first
                        List<string> drugSeqs = new List<string>();
                        string selectSql = @"
                            SELECT drug_seq 
                            FROM tbsid040_03 
                            WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @srcChart 
                              AND pat_nm = @pat_nm 
                              AND pat_jumin_no = @pat_jumin_no";

                        using (SqlCommand selectCmd = new SqlCommand(selectSql, conn, trans))
                        {
                            selectCmd.Parameters.AddWithValue("@srcChart", srcChart);
                            selectCmd.Parameters.AddWithValue("@pat_nm", patNm);
                            selectCmd.Parameters.AddWithValue("@pat_jumin_no", jumin);

                            using (SqlDataReader r = selectCmd.ExecuteReader())
                            {
                                while (r.Read())
                                {
                                    drugSeqs.Add(r[0].ToString());
                                }
                            }
                        }

                        if (drugSeqs.Count == 0)
                        {
                            trans.Rollback();
                            MessageBox.Show("이관할 조건의 처방전 데이터를 찾을 수 없습니다.", "이관 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        // 2. Update prescriptions (tbsid040_03)
                        string updateRxSql = @"
                            UPDATE tbsid040_03
                            SET chrtno = @destChart,
                                pat_nm = @destName
                            WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @srcChart 
                              AND pat_nm = @pat_nm 
                              AND pat_jumin_no = @pat_jumin_no";

                        int rxRowsAffected = 0;
                        using (SqlCommand updateRxCmd = new SqlCommand(updateRxSql, conn, trans))
                        {
                            updateRxCmd.Parameters.AddWithValue("@destChart", destChart);
                            updateRxCmd.Parameters.AddWithValue("@destName", destName);
                            updateRxCmd.Parameters.AddWithValue("@srcChart", srcChart);
                            updateRxCmd.Parameters.AddWithValue("@pat_nm", patNm);
                            updateRxCmd.Parameters.AddWithValue("@pat_jumin_no", jumin);
                            rxRowsAffected = updateRxCmd.ExecuteNonQuery();
                        }

                        // 3. Update receipts (TBSIR000_01)
                        int rcpRowsAffected = 0;
                        if (drugSeqs.Count > 0)
                        {
                            List<string> quotedSeqs = new List<string>();
                            foreach (string seq in drugSeqs)
                            {
                                quotedSeqs.Add("'" + seq.Replace("'", "''") + "'");
                            }

                            string updateRcpSql = @"
                                UPDATE TBSIR000_01
                                SET chrtno = @destChart
                                WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @srcChart 
                                  AND drug_seq IN (" + string.Join(",", quotedSeqs.ToArray()) + ")";

                            using (SqlCommand updateRcpCmd = new SqlCommand(updateRcpSql, conn, trans))
                            {
                                updateRcpCmd.Parameters.AddWithValue("@destChart", destChart);
                                updateRcpCmd.Parameters.AddWithValue("@srcChart", srcChart);
                                rcpRowsAffected = updateRcpCmd.ExecuteNonQuery();
                            }
                        }

                        // 4. Deactivate source chart if it shares same Jumin and has no rx left
                        string timeStr = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string deactivateSql = @"
                            UPDATE tbsit000_01
                            SET cusact = '0',
                                proc_dtime = @proc_dtime
                            WHERE chrtno = @srcChart
                              AND jumin_encrypt = (SELECT jumin_encrypt FROM tbsit000_01 WHERE chrtno = @destChart)
                              AND NOT EXISTS (SELECT 1 FROM tbsid040_03 WHERE COALESCE(NULLIF(LTRIM(RTRIM(chrtno)), ''), '') = @srcChart)";

                        using (SqlCommand deactCmd = new SqlCommand(deactivateSql, conn, trans))
                        {
                            deactCmd.Parameters.AddWithValue("@srcChart", srcChart);
                            deactCmd.Parameters.AddWithValue("@destChart", destChart);
                            deactCmd.Parameters.AddWithValue("@proc_dtime", timeStr);
                            deactCmd.ExecuteNonQuery();
                        }

                        trans.Commit();
                        MessageBox.Show(string.Format("이관 성공!\n\n- 처방전 변경: {0}건\n- 수납 내역 변경: {1}건\n- 차트번호 [{2}] ➔ [{3}] ({4}) 로 완료되었습니다.", 
                            rxRowsAffected, rcpRowsAffected, srcChart, destChart, destName), "이관 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        MessageBox.Show("이관 실행 중 오류가 발생하여 모든 변경 사항이 취소(Rollback)되었습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            this.Cursor = Cursors.Default;
            LoadChartDetails(srcChart);
        }

        private void BtnQuickSolve_Click(object sender, EventArgs e)
        {
            string chrtno = _txtChartNo.Text.Trim();
            if (chrtno != "0000184791") return;

            DialogResult dr = MessageBox.Show(
                "박복순-천미선 차트 혼선 건에 대하여 아래의 자동 조치를 원클릭으로 일괄 실행하시겠습니까?\n\n" +
                "1. [0000184791] 차트 마스터를 '천미선'으로 복구\n" +
                "2. [0000184791] 에 잘못 기록된 '박복순' 처방 1건을 진짜 박복순 차트인 [0000144177] 로 즉시 이동",
                "원클릭 자동 해결 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            // Find candidates
            PatientGroup pgCheon = null;
            PatientGroup pgPark = null;

            foreach (ComboItem item in _cmbRestoreCandidates.Items)
            {
                PatientGroup pg = (PatientGroup)item.Value;
                if (pg.Name == "천미선") pgCheon = pg;
                if (pg.Name == "박복순") pgPark = pg;
            }

            if (pgCheon == null || pgPark == null)
            {
                MessageBox.Show("원클릭 해결에 필요한 처방 분포 데이터(천미선/박복순)가 조회되지 않습니다.", "실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _mainForm.RestoreMaster(chrtno, pgCheon);

            if (_isDemo)
            {
                MoveRxDemo(chrtno, pgPark.Name, pgPark.Jumin, "0000144177");
            }
            else
            {
                MoveRxProduction(chrtno, pgPark.Name, pgPark.Jumin, "0000144177");
            }

            MessageBox.Show("박복순-천미선 차트 혼선 원클릭 자동 조치 완료!", "해결 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void LoadChartDetails(string chrtno)
        {
            if (_isDemo)
            {
                LoadChartDemo(chrtno);
            }
            else
            {
                LoadChartProduction(chrtno);
            }
        }

        private void BtnCreateNewChart_Click(object sender, EventArgs e)
        {
            string newChartNo = GenerateNewChartNo();
            if (!string.IsNullOrEmpty(newChartNo))
            {
                _txtCustomDestChart.Text = newChartNo;
                MessageBox.Show(string.Format("미사용 중인 가장 높은 신규 차트번호 [{0}]을(를) 생성했습니다.\n\n이 번호로 이송을 실행하시면 신규 고객으로 자동 등록됩니다.", newChartNo), "새 차트번호 발급", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string GenerateNewChartNo()
        {
            long maxNum = 0;
            if (_isDemo)
            {
                foreach (var c in _mainForm._mockCustList)
                {
                    long val;
                    if (long.TryParse(c.ChrtNo, out val))
                    {
                        if (val > maxNum) maxNum = val;
                    }
                }
                foreach (var rx in _mainForm._mockRxList)
                {
                    long val;
                    if (long.TryParse(rx.ChrtNo, out val))
                    {
                        if (val > maxNum) maxNum = val;
                    }
                }
            }
            else
            {
                string connStr = _mainForm.BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT MAX(chrtno) FROM tbsit000_01";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            var res = cmd.ExecuteScalar();
                            if (res != DBNull.Value && res != null)
                            {
                                long val;
                                if (long.TryParse(res.ToString(), out val))
                                {
                                    maxNum = val;
                                }
                            }
                        }
                        string sqlRx = "SELECT MAX(chrtno) FROM tbsid040_03";
                        using (SqlCommand cmd = new SqlCommand(sqlRx, conn))
                        {
                            var res = cmd.ExecuteScalar();
                            if (res != DBNull.Value && res != null)
                            {
                                long val;
                                if (long.TryParse(res.ToString(), out val))
                                {
                                    if (val > maxNum) maxNum = val;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("신규 차트번호 생성 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "";
                }
            }

            if (maxNum == 0) maxNum = 100000;
            
            long candidate = maxNum + 1;
            while (true)
            {
                string chrt = candidate.ToString("D10");
                if (CheckChartNoExists(chrt))
                {
                    candidate++;
                }
                else
                {
                    return chrt;
                }
            }
        }

        private bool CheckChartNoExists(string chrtno)
        {
            if (_isDemo)
            {
                return _mainForm._mockCustList.Exists(c => c.ChrtNo == chrtno) || 
                       _mainForm._mockRxList.Exists(rx => rx.ChrtNo == chrtno);
            }
            else
            {
                string connStr = _mainForm.BuildConnectionString(false);
                try
                {
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string sql = "SELECT COUNT(*) FROM tbsit000_01 WHERE chrtno = @chrtno";
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.AddWithValue("@chrtno", chrtno);
                            int count = Convert.ToInt32(cmd.ExecuteScalar());
                            if (count > 0) return true;
                        }
                        string sqlRx = "SELECT COUNT(*) FROM tbsid040_03 WHERE chrtno = @chrtno";
                        using (SqlCommand cmd = new SqlCommand(sqlRx, conn))
                        {
                            cmd.Parameters.AddWithValue("@chrtno", chrtno);
                            int count = Convert.ToInt32(cmd.ExecuteScalar());
                            if (count > 0) return true;
                        }
                    }
                }
                catch (Exception)
                {
                    return true;
                }
                return false;
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (_dgvDuplicate.Rows.Count == 0)
            {
                MessageBox.Show("내보낼 데이터가 없습니다. 먼저 스캔을 실행해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                sfd.FileName = string.Format("개명_및_중복_차트_목록_{0}.csv", DateTime.Now.ToString("yyyyMMdd"));
                sfd.Title = "개명 및 중복 차트 목록 내보내기";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToCsv(_dgvDuplicate, sfd.FileName);
                        MessageBox.Show("성공적으로 저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("파일 저장 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnDeleteEmptyDuplicates_Click(object sender, EventArgs e)
        {
            if (_isDemo)
            {
                DeleteEmptyDuplicatesDemo();
            }
            else
            {
                DeleteEmptyDuplicatesProduction();
            }
        }

        private void DeleteEmptyDuplicatesDemo()
        {
            int deletedCount = 0;
            var groups = _mainForm._mockCustList.Where(c => c.CusAct == "1" && c.PatJuminNo.Replace("-","").Length == 13).GroupBy(c => c.JuminEncrypt).ToList();
            foreach (var g in groups)
            {
                if (g.Count() > 1)
                {
                    var sorted = g.Select(c => new {
                        Cust = c,
                        RxCount = _mainForm._mockRxList.Count(r => r.ChrtNo == c.ChrtNo)
                    })
                    .OrderByDescending(x => x.RxCount)
                    .ThenBy(x => x.Cust.ChrtNo)
                    .ToList();
                    
                    for (int i = 1; i < sorted.Count; i++)
                    {
                        if (sorted[i].RxCount == 0)
                        {
                            _mainForm._mockCustList.Remove(sorted[i].Cust);
                            deletedCount++;
                        }
                    }
                }
            }
            
            MessageBox.Show(string.Format("[데모] 처방 내역이 없는 중복 차트 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            BtnScanMixed_Click(null, null);
        }

        private void DeleteEmptyDuplicatesProduction()
        {
            string connStr = _mainForm.BuildConnectionString(false);
            int candidateCount = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string countSql = @"
                        WITH DuplicateCharts AS (
                            SELECT m.chrtno,
                                   (SELECT COUNT(*) FROM tbsid040_03 r WHERE r.chrtno = m.chrtno) as rx_count,
                                   ROW_NUMBER() OVER (
                                       PARTITION BY m.jumin_encrypt 
                                       ORDER BY (SELECT COUNT(*) FROM tbsid040_03 r WHERE r.chrtno = m.chrtno) DESC, m.chrtno ASC
                                   ) as rn
                            FROM tbsit000_01 m
                            WHERE m.cusact = '1' 
                              AND m.jumin_no LIKE '[0-9]%'
                        )
                        SELECT COUNT(*) 
                        FROM DuplicateCharts 
                        WHERE rn > 1 AND rx_count = 0;";
                    
                    using (SqlCommand cmd = new SqlCommand(countSql, conn))
                    {
                        candidateCount = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("중복 차트 조회 실패:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if (candidateCount == 0)
            {
                MessageBox.Show("삭제 대상인 '처방 내역이 없는 중복 차트'가 존재하지 않습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            DialogResult dr = MessageBox.Show(
                string.Format("처방 내역이 없으면서 다른 활성 차트가 존재하는 중복 차트 {0}건을 정말로 영구 삭제하시겠습니까?\n\n" +
                              "※ 주민번호당 처방 내역이 가장 많은 차트(또는 최초 생성 차트) 1건은 삭제되지 않고 안전하게 보존됩니다.", candidateCount),
                "중복 차트 일괄 삭제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );
            
            if (dr != DialogResult.Yes) return;
            
            this.Cursor = Cursors.WaitCursor;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int deletedCount = 0;
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        string deleteSql = @"
                            WITH DuplicateCharts AS (
                                SELECT m.chrtno,
                                       (SELECT COUNT(*) FROM tbsid040_03 r WHERE r.chrtno = m.chrtno) as rx_count,
                                       ROW_NUMBER() OVER (
                                           PARTITION BY m.jumin_encrypt 
                                           ORDER BY (SELECT COUNT(*) FROM tbsid040_03 r WHERE r.chrtno = m.chrtno) DESC, m.chrtno ASC
                                       ) as rn
                                FROM tbsit000_01 m
                                WHERE m.cusact = '1' 
                                  AND m.jumin_no LIKE '[0-9]%'
                            )
                            DELETE FROM tbsit000_01
                            WHERE chrtno IN (
                                SELECT chrtno 
                                FROM DuplicateCharts 
                                WHERE rn > 1 AND rx_count = 0
                            );";
                        
                        using (SqlCommand cmd = new SqlCommand(deleteSql, conn))
                        {
                            cmd.CommandTimeout = 300;
                            deletedCount = cmd.ExecuteNonQuery();
                        }
                    }
                    
                    this.BeginInvoke((Action)(() =>
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(string.Format("처방 내역이 없는 중복 차트 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "삭제 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        BtnScanMixed_Click(null, null);
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("중복 차트 삭제 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void BtnDeleteGhostCharts_Click(object sender, EventArgs e)
        {
            if (_isDemo)
            {
                DeleteGhostChartsDemo();
            }
            else
            {
                DeleteGhostChartsProduction();
            }
        }

        private void DeleteGhostChartsDemo()
        {
            int deletedCount = _mainForm._mockCustList.RemoveAll(c => c.CusAct == "1" && (string.IsNullOrEmpty(c.PatNm) || c.PatNm.Trim() == ""));
            MessageBox.Show(string.Format("[데모] 이름이 없는 유령 환자 차트 {0}건이 정상적으로 삭제되었습니다.", deletedCount), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            BtnScanMixed_Click(null, null);
        }

        private void DeleteGhostChartsProduction()
        {
            string connStr = _mainForm.BuildConnectionString(false);
            int candidateCount = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    string countSql = @"
                        SELECT COUNT(*) 
                        FROM tbsit000_01
                        WHERE cusact = '1'
                          AND (pat_nm IS NULL 
                               OR LTRIM(RTRIM(pat_nm)) = '' 
                               OR jumin_encrypt = 'kK9LhrP2HFOC+IcfVnNMTg==01');";
                    
                    using (SqlCommand cmd = new SqlCommand(countSql, conn))
                    {
                        candidateCount = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("유령 차트 조회 중 오류 발생:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (candidateCount == 0)
            {
                MessageBox.Show("삭제 대상인 '이름과 주민번호가 없는 유령 차트'가 존재하지 않습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult dr = MessageBox.Show(
                string.Format("성명과 주민번호가 없는 유령 환자 차트 {0}건을 정말로 영구 일괄 삭제하시겠습니까?\n\n" +
                              "※ 주의: 실제 처방 내역이 없는 안전한 유령 차트만 삭제되며, 삭제 후에는 복구할 수 없습니다.", candidateCount),
                "유령 차트 일괄 삭제 확인",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (dr != DialogResult.Yes) return;

            this.Cursor = Cursors.WaitCursor;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int deletedCount = 0;
                    using (SqlConnection conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        using (SqlTransaction trans = conn.BeginTransaction())
                        {
                            try
                            {
                                string deleteSql = @"
                                    DELETE FROM tbsit000_01
                                    WHERE cusact = '1'
                                      AND (pat_nm IS NULL OR LTRIM(RTRIM(pat_nm)) = '' OR jumin_encrypt = 'kK9LhrP2HFOC+IcfVnNMTg==01')
                                      AND chrtno NOT IN (
                                          SELECT DISTINCT CHRTNO 
                                          FROM tbsid040_03 
                                          WHERE CHRTNO IS NOT NULL AND LTRIM(RTRIM(CHRTNO)) <> ''
                                      );";

                                using (SqlCommand cmd = new SqlCommand(deleteSql, conn, trans))
                                {
                                    deletedCount = cmd.ExecuteNonQuery();
                                }
                                trans.Commit();
                            }
                            catch (Exception)
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }

                    this.BeginInvoke((Action)(() =>
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(string.Format("유령 환자 차트 {0}건이 정상적으로 영구 삭제되었습니다.", deletedCount), "삭제 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        BtnScanMixed_Click(null, null);
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("유령 차트 일괄 삭제 중 오류가 발생했습니다:\n" + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
            });
        }

        private void ExportToCsv(DataGridView dgv, string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, new System.Text.UTF8Encoding(true)))
            {
                // Write Headers
                List<string> headers = new List<string>();
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    headers.Add(EscapeCsv(col.HeaderText));
                }
                sw.WriteLine(string.Join(",", headers.ToArray()));

                // Write Rows
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    List<string> cells = new List<string>();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cells.Add(EscapeCsv(cell.Value != null ? cell.Value.ToString() : ""));
                    }
                    sw.WriteLine(string.Join(",", cells.ToArray()));
                }
            }
        }

        private string EscapeCsv(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";

            bool needsQuotes = val.Contains(",") || val.Contains("\"") || val.Contains("\n") || val.Contains("\r");
            if (needsQuotes)
            {
                return "\"" + val.Replace("\"", "\"\"") + "\"";
            }
            return val;
        }
    }
}