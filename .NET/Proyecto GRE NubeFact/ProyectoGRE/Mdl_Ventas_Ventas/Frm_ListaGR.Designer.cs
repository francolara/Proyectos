
namespace ProyectoGRE
{
    partial class Frm_ListaGR
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_ListaGR));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CmdEnvSunat = new System.Windows.Forms.Button();
            this.Dgv_Guias = new System.Windows.Forms.DataGridView();
            this.Tipo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NumOriEfc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SERIE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.numeroori = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NumDoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FechaDoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AL1_NOMCLIPRO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AL1_TIPIGV = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AL1_TIPMON = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AL1_TOTVTA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.NTFNB = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ejecutarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.Dgv_Guias)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(5, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(158, 34);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // CmdEnvSunat
            // 
            this.CmdEnvSunat.Location = new System.Drawing.Point(57, 142);
            this.CmdEnvSunat.Name = "CmdEnvSunat";
            this.CmdEnvSunat.Size = new System.Drawing.Size(155, 33);
            this.CmdEnvSunat.TabIndex = 3;
            this.CmdEnvSunat.Text = "Enviar y Consultar";
            this.CmdEnvSunat.UseVisualStyleBackColor = true;
            this.CmdEnvSunat.Click += new System.EventHandler(this.CmdEnvSunat_Click);
            // 
            // Dgv_Guias
            // 
            this.Dgv_Guias.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Dgv_Guias.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Tipo,
            this.NumOriEfc,
            this.SERIE,
            this.numeroori,
            this.NumDoc,
            this.FechaDoc,
            this.AL1_NOMCLIPRO,
            this.AL1_TIPIGV,
            this.AL1_TIPMON,
            this.AL1_TOTVTA});
            this.Dgv_Guias.Location = new System.Drawing.Point(371, 135);
            this.Dgv_Guias.Name = "Dgv_Guias";
            this.Dgv_Guias.RowHeadersWidth = 25;
            this.Dgv_Guias.RowTemplate.Height = 25;
            this.Dgv_Guias.Size = new System.Drawing.Size(307, 240);
            this.Dgv_Guias.TabIndex = 1;
            // 
            // Tipo
            // 
            this.Tipo.DataPropertyName = "Tipo";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft;
            this.Tipo.DefaultCellStyle = dataGridViewCellStyle1;
            this.Tipo.HeaderText = "Tipo";
            this.Tipo.Name = "Tipo";
            this.Tipo.ReadOnly = true;
            this.Tipo.Width = 40;
            // 
            // NumOriEfc
            // 
            this.NumOriEfc.DataPropertyName = "NumOriEfc";
            this.NumOriEfc.HeaderText = "NumOriEfc";
            this.NumOriEfc.Name = "NumOriEfc";
            this.NumOriEfc.Visible = false;
            // 
            // SERIE
            // 
            this.SERIE.DataPropertyName = "SERIE";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.SERIE.DefaultCellStyle = dataGridViewCellStyle2;
            this.SERIE.HeaderText = "Serie";
            this.SERIE.Name = "SERIE";
            this.SERIE.ReadOnly = true;
            this.SERIE.Visible = false;
            this.SERIE.Width = 50;
            // 
            // numeroori
            // 
            this.numeroori.DataPropertyName = "numeroori";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.numeroori.DefaultCellStyle = dataGridViewCellStyle3;
            this.numeroori.HeaderText = "Numero";
            this.numeroori.Name = "numeroori";
            this.numeroori.ReadOnly = true;
            this.numeroori.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.numeroori.Width = 115;
            // 
            // NumDoc
            // 
            this.NumDoc.DataPropertyName = "NumDoc";
            this.NumDoc.HeaderText = "NumDoc";
            this.NumDoc.Name = "NumDoc";
            this.NumDoc.Visible = false;
            // 
            // FechaDoc
            // 
            this.FechaDoc.DataPropertyName = "FechaDoc";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.FechaDoc.DefaultCellStyle = dataGridViewCellStyle4;
            this.FechaDoc.HeaderText = "Fecha";
            this.FechaDoc.Name = "FechaDoc";
            this.FechaDoc.ReadOnly = true;
            // 
            // AL1_NOMCLIPRO
            // 
            this.AL1_NOMCLIPRO.DataPropertyName = "AL1_NOMCLIPRO";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.AL1_NOMCLIPRO.DefaultCellStyle = dataGridViewCellStyle5;
            this.AL1_NOMCLIPRO.HeaderText = "Razón Social";
            this.AL1_NOMCLIPRO.Name = "AL1_NOMCLIPRO";
            this.AL1_NOMCLIPRO.ReadOnly = true;
            this.AL1_NOMCLIPRO.Visible = false;
            this.AL1_NOMCLIPRO.Width = 250;
            // 
            // AL1_TIPIGV
            // 
            this.AL1_TIPIGV.DataPropertyName = "AL1_TIPIGV";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.AL1_TIPIGV.DefaultCellStyle = dataGridViewCellStyle6;
            this.AL1_TIPIGV.HeaderText = "Tipo IGV";
            this.AL1_TIPIGV.Name = "AL1_TIPIGV";
            this.AL1_TIPIGV.ReadOnly = true;
            this.AL1_TIPIGV.Visible = false;
            this.AL1_TIPIGV.Width = 80;
            // 
            // AL1_TIPMON
            // 
            this.AL1_TIPMON.DataPropertyName = "AL1_TIPMON";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter;
            this.AL1_TIPMON.DefaultCellStyle = dataGridViewCellStyle7;
            this.AL1_TIPMON.HeaderText = "Moneda";
            this.AL1_TIPMON.Name = "AL1_TIPMON";
            this.AL1_TIPMON.ReadOnly = true;
            this.AL1_TIPMON.Visible = false;
            this.AL1_TIPMON.Width = 70;
            // 
            // AL1_TOTVTA
            // 
            this.AL1_TOTVTA.DataPropertyName = "AL1_TOTVTA";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight;
            this.AL1_TOTVTA.DefaultCellStyle = dataGridViewCellStyle8;
            this.AL1_TOTVTA.HeaderText = "Total";
            this.AL1_TOTVTA.Name = "AL1_TOTVTA";
            this.AL1_TOTVTA.ReadOnly = true;
            this.AL1_TOTVTA.Visible = false;
            this.AL1_TOTVTA.Width = 70;
            // 
            // timer1
            // 
            this.timer1.Interval = 15000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // NTFNB
            // 
            this.NTFNB.ContextMenuStrip = this.contextMenuStrip1;
            this.NTFNB.Icon = ((System.Drawing.Icon)(resources.GetObject("NTFNB.Icon")));
            this.NTFNB.Text = "Servicio Sunat";
            this.NTFNB.Visible = true;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ejecutarToolStripMenuItem,
            this.salirToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(117, 48);
            // 
            // ejecutarToolStripMenuItem
            // 
            this.ejecutarToolStripMenuItem.Name = "ejecutarToolStripMenuItem";
            this.ejecutarToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.ejecutarToolStripMenuItem.Text = "Ejecutar";
            this.ejecutarToolStripMenuItem.Click += new System.EventHandler(this.ejecutarToolStripMenuItem_Click);
            // 
            // salirToolStripMenuItem
            // 
            this.salirToolStripMenuItem.Name = "salirToolStripMenuItem";
            this.salirToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.salirToolStripMenuItem.Text = "Salir";
            this.salirToolStripMenuItem.Click += new System.EventHandler(this.salirToolStripMenuItem_Click_1);
            // 
            // Frm_ListaGR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(243, 175);
            this.Controls.Add(this.CmdEnvSunat);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Dgv_Guias);
            this.MaximizeBox = false;
            this.Name = "Frm_ListaGR";
            this.Text = "Envio Sunat";
            this.Load += new System.EventHandler(this.Frm_ListaGR_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Dgv_Guias)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView Dgv_Guias;
        private System.Windows.Forms.Button CmdEnvSunat;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tipo;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumOriEfc;
        private System.Windows.Forms.DataGridViewTextBoxColumn SERIE;
        private System.Windows.Forms.DataGridViewTextBoxColumn numeroori;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumDoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn FechaDoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn AL1_NOMCLIPRO;
        private System.Windows.Forms.DataGridViewTextBoxColumn AL1_TIPIGV;
        private System.Windows.Forms.DataGridViewTextBoxColumn AL1_TIPMON;
        private System.Windows.Forms.DataGridViewTextBoxColumn AL1_TOTVTA;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.NotifyIcon NTFNB;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ejecutarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salirToolStripMenuItem;
    }
}

