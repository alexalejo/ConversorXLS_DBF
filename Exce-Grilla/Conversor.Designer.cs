namespace Exce_Grilla
{
    partial class Conversor
    {
        /// <summary>
        /// Variable del diseñador requerida.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén utilizando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben eliminar; false en caso contrario, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.cmbHojas = new System.Windows.Forms.ComboBox();
            this.ofdArchivo = new System.Windows.Forms.OpenFileDialog();
            this.brProgreso = new System.Windows.Forms.ProgressBar();
            this.btnExaminar = new System.Windows.Forms.Button();
            this.btnComprobar = new System.Windows.Forms.Button();
            this.btnprocesar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSalir = new System.Windows.Forms.Button();
            this.lRuta = new System.Windows.Forms.Label();
            this.dgDatos = new System.Windows.Forms.DataGridView();
            this.dsDocumentos = new Exce_Grilla.dsDocumentos();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgDatos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsDocumentos)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbHojas
            // 
            this.cmbHojas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbHojas.Location = new System.Drawing.Point(130, 48);
            this.cmbHojas.Name = "cmbHojas";
            this.cmbHojas.Size = new System.Drawing.Size(160, 21);
            this.cmbHojas.TabIndex = 1;
            this.cmbHojas.SelectedIndexChanged += new System.EventHandler(this.cmbHojas_SelectedIndexChanged);
            // 
            // ofdArchivo
            // 
            this.ofdArchivo.FileName = "openFileDialog1";
            // 
            // brProgreso
            // 
            this.brProgreso.Location = new System.Drawing.Point(130, 80);
            this.brProgreso.Name = "brProgreso";
            this.brProgreso.Size = new System.Drawing.Size(160, 19);
            this.brProgreso.TabIndex = 3;
            // 
            // btnExaminar
            // 
            this.btnExaminar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExaminar.Location = new System.Drawing.Point(661, 19);
            this.btnExaminar.Name = "btnExaminar";
            this.btnExaminar.Size = new System.Drawing.Size(93, 23);
            this.btnExaminar.TabIndex = 4;
            this.btnExaminar.Text = "Buscar Archivo";
            this.btnExaminar.UseVisualStyleBackColor = true;
            this.btnExaminar.Click += new System.EventHandler(this.btnExaminar_Click);
            // 
            // btnComprobar
            // 
            this.btnComprobar.Location = new System.Drawing.Point(661, 49);
            this.btnComprobar.Name = "btnComprobar";
            this.btnComprobar.Size = new System.Drawing.Size(93, 23);
            this.btnComprobar.TabIndex = 5;
            this.btnComprobar.Text = "Comprobar Hoja";
            this.btnComprobar.UseVisualStyleBackColor = true;
            this.btnComprobar.Click += new System.EventHandler(this.btnComprobar_Click);
            // 
            // btnprocesar
            // 
            this.btnprocesar.Location = new System.Drawing.Point(661, 80);
            this.btnprocesar.Name = "btnprocesar";
            this.btnprocesar.Size = new System.Drawing.Size(93, 23);
            this.btnprocesar.TabIndex = 6;
            this.btnprocesar.Text = "Procesar Datos";
            this.btnprocesar.UseVisualStyleBackColor = true;
            this.btnprocesar.Click += new System.EventHandler(this.btnprocesar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(40, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 15);
            this.label1.TabIndex = 7;
            this.label1.Text = "Archivo Excel:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label2.Location = new System.Drawing.Point(24, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Nombre de la Hoja:";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(460, 80);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(183, 17);
            this.checkBox1.TabIndex = 9;
            this.checkBox1.Text = "Eliminar Artículos con Stock cero";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label3.Location = new System.Drawing.Point(64, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Progreso...";
            // 
            // btnSalir
            // 
            this.btnSalir.Location = new System.Drawing.Point(671, 404);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(93, 23);
            this.btnSalir.TabIndex = 12;
            this.btnSalir.Text = "Salir";
            this.btnSalir.UseVisualStyleBackColor = true;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
            // 
            // lRuta
            // 
            this.lRuta.BackColor = System.Drawing.Color.White;
            this.lRuta.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lRuta.Enabled = false;
            this.lRuta.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lRuta.Location = new System.Drawing.Point(130, 21);
            this.lRuta.Name = "lRuta";
            this.lRuta.Size = new System.Drawing.Size(513, 19);
            this.lRuta.TabIndex = 13;
            // 
            // dgDatos
            // 
            this.dgDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgDatos.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgDatos.Location = new System.Drawing.Point(12, 108);
            this.dgDatos.Name = "dgDatos";
            this.dgDatos.Size = new System.Drawing.Size(752, 290);
            this.dgDatos.TabIndex = 14;
            // 
            // dsDocumentos
            // 
            this.dsDocumentos.DataSetName = "dsDocumentos";
            this.dsDocumentos.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(325, 46);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 15;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(325, 74);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Conversor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(776, 430);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dgDatos);
            this.Controls.Add(this.lRuta);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnprocesar);
            this.Controls.Add(this.btnComprobar);
            this.Controls.Add(this.btnExaminar);
            this.Controls.Add(this.brProgreso);
            this.Controls.Add(this.cmbHojas);
            this.Name = "Conversor";
            this.Text = "Mantenimiento de Artículos (Factura Plus)";
            ((System.ComponentModel.ISupportInitialize)(this.dgDatos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsDocumentos)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbHojas;
        private System.Windows.Forms.OpenFileDialog ofdArchivo;
        private System.Windows.Forms.ProgressBar brProgreso;
        private System.Windows.Forms.Button btnExaminar;
        private System.Windows.Forms.Button btnComprobar;
        private System.Windows.Forms.Button btnprocesar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSalir;
        private System.Windows.Forms.Label lRuta;
        private dsDocumentos dsDocumentos;
        private System.Windows.Forms.DataGridView dgDatos;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

