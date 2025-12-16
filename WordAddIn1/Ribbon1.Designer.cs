namespace WordAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button8 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group5.SuspendLayout();
            this.group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Label = "리앤권 작업도구";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button7);
            this.group1.Label = "도면 작업";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "도면 삽입";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImageInserter_Click);
            // 
            // button7
            // 
            this.button7.Label = "도면 점검";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DrawingValidatorClick);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button11);
            this.group2.Label = "청구항 작업";
            this.group2.Name = "group2";
            // 
            // button4
            // 
            this.button4.Label = "청구항 구성 분할";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextSpliter_Click);
            // 
            // button11
            // 
            this.button11.Label = "청구항 트리 작성";
            this.button11.Name = "button11";
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerateClaimTable_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button6);
            this.group3.Items.Add(this.button10);
            this.group3.Items.Add(this.button5);
            this.group3.Label = "삭제 도구";
            this.group3.Name = "group3";
            // 
            // button6
            // 
            this.button6.Label = "도면 부호 삭제";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DrawingReferenceRemover_Click);
            // 
            // button10
            // 
            this.button10.Label = "메모 삭제";
            this.button10.Name = "button10";
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CommentRemover_Click);
            // 
            // button5
            // 
            this.button5.Label = "명세서 초안 마무리";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PrepareSubmissionFile_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.button3);
            this.group5.Items.Add(this.button9);
            this.group5.Label = "변환 도구";
            this.group5.Name = "group5";
            // 
            // button3
            // 
            this.button3.Label = "문서 작성자 변경";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AuthorChanger_Click);
            // 
            // button9
            // 
            this.button9.Label = "Paraphrsing";
            this.button9.Name = "button9";
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClaimParaphraser_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button8);
            this.group4.Label = "PCT 명세서 작업";
            this.group4.Name = "group4";
            // 
            // button8
            // 
            this.button8.Label = "형식 변환";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PCTTransformer_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.button2);
            this.group6.Items.Add(this.button12);
            this.group6.Label = "비교 도구";
            this.group6.Name = "group6";
            // 
            // button2
            // 
            this.button2.Label = "좌우 청구항 비교";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "표를 생성하여 왼쪽 셀에 원본 청구항을 붙여 넣고, 오른쪽 셀에 수정된 청구항을 붙여 넣은 후, 표 안에서 버튼을 클릭해주세요";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableComparer_Click);
            // 
            // button12
            // 
            this.button12.Label = "현재 문서와 비교";
            this.button12.Name = "button12";
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DocumentComparer_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.button13);
            this.group7.Items.Add(this.button14);
            this.group7.Label = "서식 도구";
            this.group7.Name = "group7";
            // 
            // button13
            // 
            this.button13.Label = "KR명세서 서식 적용";
            this.button13.Name = "button13";
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyPatentFormat_Click);
            // 
            // button14
            // 
            this.button14.Label = "추적변경된 서식만 적용";
            this.button14.Name = "button14";
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
