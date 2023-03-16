namespace Graficador___v3
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
            this.GraficadorFrecuenciaGanancia = new System.Windows.Forms.Button();
            this.GráficaFrecuenciaFase = new System.Windows.Forms.Button();
            this.BorraGráficas = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // GraficadorFrecuenciaGanancia
            // 
            this.GraficadorFrecuenciaGanancia.Dock = System.Windows.Forms.DockStyle.Top;
            this.GraficadorFrecuenciaGanancia.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.GraficadorFrecuenciaGanancia.Location = new System.Drawing.Point(0, 0);
            this.GraficadorFrecuenciaGanancia.Name = "GraficadorFrecuenciaGanancia";
            this.GraficadorFrecuenciaGanancia.Size = new System.Drawing.Size(800, 153);
            this.GraficadorFrecuenciaGanancia.TabIndex = 0;
            this.GraficadorFrecuenciaGanancia.Tag = "GrafFrecGan";
            this.GraficadorFrecuenciaGanancia.Text = "Gráfica Frecuencia-Ganancia";
            this.GraficadorFrecuenciaGanancia.UseVisualStyleBackColor = true;
            this.GraficadorFrecuenciaGanancia.Click += new System.EventHandler(this.GraficadorFrecuenciaGanancia_Click);
            // 
            // GráficaFrecuenciaFase
            // 
            this.GráficaFrecuenciaFase.Dock = System.Windows.Forms.DockStyle.Top;
            this.GráficaFrecuenciaFase.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.GráficaFrecuenciaFase.Location = new System.Drawing.Point(0, 153);
            this.GráficaFrecuenciaFase.Name = "GráficaFrecuenciaFase";
            this.GráficaFrecuenciaFase.Size = new System.Drawing.Size(800, 153);
            this.GráficaFrecuenciaFase.TabIndex = 1;
            this.GráficaFrecuenciaFase.Tag = "GrafFrecFase";
            this.GráficaFrecuenciaFase.Text = "Gráfica Frecuencia-Fase";
            this.GráficaFrecuenciaFase.UseVisualStyleBackColor = true;
            this.GráficaFrecuenciaFase.Click += new System.EventHandler(this.GráficaFrecuenciaFase_Click);
            // 
            // BorraGráficas
            // 
            this.BorraGráficas.Dock = System.Windows.Forms.DockStyle.Top;
            this.BorraGráficas.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.BorraGráficas.Location = new System.Drawing.Point(0, 306);
            this.BorraGráficas.Name = "BorraGráficas";
            this.BorraGráficas.Size = new System.Drawing.Size(800, 153);
            this.BorraGráficas.TabIndex = 2;
            this.BorraGráficas.Tag = "BorraGráficas";
            this.BorraGráficas.Text = "Borrar Gráficas";
            this.BorraGráficas.UseVisualStyleBackColor = true;
            this.BorraGráficas.Click += new System.EventHandler(this.BorraGráficas_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.BorraGráficas);
            this.Controls.Add(this.GráficaFrecuenciaFase);
            this.Controls.Add(this.GraficadorFrecuenciaGanancia);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private Button GraficadorFrecuenciaGanancia;
        private Button GráficaFrecuenciaFase;
        private Button BorraGráficas;
    }
}