namespace Proyecto_Grilla
{
    partial class Form1
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
            btnExcel = new Button();
            dataGridView1 = new DataGridView();
            Nombre = new DataGridViewTextBoxColumn();
            Apellido = new DataGridViewTextBoxColumn();
            Telefono = new DataGridViewTextBoxColumn();
            Email = new DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // btnExcel
            // 
            btnExcel.Location = new Point(309, 242);
            btnExcel.Name = "btnExcel";
            btnExcel.Size = new Size(117, 23);
            btnExcel.TabIndex = 0;
            btnExcel.Text = "Exportar a Excel";
            btnExcel.UseVisualStyleBackColor = true;
            btnExcel.Click += btnExcel_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { Nombre, Apellido, Telefono, Email });
            dataGridView1.Location = new Point(32, 12);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(565, 150);
            dataGridView1.TabIndex = 0;
            // 
            // Nombre
            // 
            Nombre.HeaderText = "Nombre";
            Nombre.Name = "Nombre";
            // 
            // Apellido
            // 
            Apellido.HeaderText = "Apellido";
            Apellido.Name = "Apellido";
            // 
            // Telefono
            // 
            Telefono.HeaderText = "Telefono";
            Telefono.Name = "Telefono";
            // 
            // Email
            // 
            Email.HeaderText = "Email";
            Email.Name = "Email";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(dataGridView1);
            Controls.Add(btnExcel);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Button btnExcel;
        private DataGridView dataGridView1;
        private DataGridViewTextBoxColumn Nombre;
        private DataGridViewTextBoxColumn Apellido;
        private DataGridViewTextBoxColumn Telefono;
        private DataGridViewTextBoxColumn Email;
    }
}
