using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Desafio2 {
    public partial class Form1 : Form {
        private Form2 form2;
        private Form3 form3;
        private Form4 form4;

        public Form1() {
            InitializeComponent();
            
            SetFormProperties(this, 1920, 200, 0.0, 0.0, 0);
        }

        private void SetFormProperties(Form form, int width, int height, double xPercent, double yPercent, int yOffset) {
            
            var screenWidth = Screen.PrimaryScreen.Bounds.Width;
            var screenHeight = Screen.PrimaryScreen.Bounds.Height;
           
            int xPos = (int)(screenWidth * xPercent);
            int yPos = (int)(screenHeight * yPercent) + yOffset; 

            form.Size = new System.Drawing.Size(width, height);
            form.MinimumSize = new System.Drawing.Size(width, height);
            form.MaximumSize = new System.Drawing.Size(width, height);
            form.StartPosition = FormStartPosition.Manual; 
            form.Location = new System.Drawing.Point(xPos, yPos);
        }

        private async Task CloseFormsAsync(params Form[] forms) {
            var tasks = forms.Select(form => Task.Run(() => form?.Close())).ToArray();
            await Task.WhenAll(tasks);
        }

        private void OpenForm2() {
            try {
                
                form3?.Close();
                form4?.Close();

                
                if (form2 == null || form2.IsDisposed) {
                    form2 = new Form2();
                    form2.FormClosed += (s, args) => form2 = null;
                    SetFormProperties(form2, 1920, 832, 0.0, 0.0, 200);
                    form2.Show();
                }
                else {
                    form2.BringToFront();
                    form2.Focus();
                }
            }
            catch (Exception ex) {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}");
            }
        }

        private void OpenForm3() {
            try {
                
                form2?.Close();
                form4?.Close();

                
                if (form3 == null || form3.IsDisposed) {
                    form3 = new Form3();
                    form3.FormClosed += (s, args) => form3 = null; 
                    SetFormProperties(form3, 1920, 832, 0.0, 0.0, 200); 
                    form3.Show();
                }
                else {
                    form3.BringToFront();
                    form3.Focus();
                }
            }
            catch (Exception ex) {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}");
            }
        }

        private void OpenForm4() {
            try {
               
                form2?.Close();
                form3?.Close();

                
                if (form4 == null || form4.IsDisposed) {
                    form4 = new Form4();
                    form4.FormClosed += (s, args) => form4 = null; 
                    SetFormProperties(form4, 1920, 832, 0.0, 0.0, 200); 
                    form4.Show();
                }
                else {
                    form4.BringToFront();
                    form4.Focus();
                }
            }
            catch (Exception ex) {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}");
            }
        }

        private void button1_Click(object sender, EventArgs e) {
            OpenForm2();
        }

        private void button2_Click(object sender, EventArgs e) {
            OpenForm3();
        }

        private void button3_Click(object sender, EventArgs e) {
            OpenForm4();
        }

    }
}