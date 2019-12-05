using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
namespace Chuong_Trinh_Soan_Thao_Van_Ban
{
    public partial class frmworkpad : Form
    {
        private string duongdan = "";   //Tạo đường dẫn lưu hoặc mở tập tin.
        private int kiemTraSave = 0;    //0 tức là chưa lưu. 1 là đã lưu.
        private SaveFileDialog save;    //Tạo Sự kiện Dialog save để hiện thị cửa sổ save có sẵn trong hệ thống.
        private OpenFileDialog open;    //Tạo sự kiện Dialog open để hiển thị cửa sổ open có săn trong hệ thống.
        private ColorDialog color;      //Tạo Sự kiện Dialog color để hiện thị màu chữ có sẵn trong hệ thống.
        private FontDialog dialogFonts; //Tạo Sự kiện Dialog Fonts để hiện thị màu chữ có sẵn trong hệ thống.
        private int check = 0;          //Tạo biến check để kiểm tra cho các sự kiện in đậm, in nghiêng, gạch chân, canh trái, phải, giữa...


        frmtimkiem fsearch;

        public frmworkpad()
        {
            InitializeComponent();
        }
        private void frmworkpad_Load(object sender, EventArgs e)
        {
            loadFonts();
            //toolStrip_mau.BackColor = 
        }
       
       
        #region Các Phương Thức Làm Việc
        public void dinhDangFonts()
        {
            float fsize = 10;
            if (tsbcbfontsize.SelectedIndex != -1)
            {
                fsize = (float)float.Parse(tsbcbfontsize.SelectedItem.ToString());
            }
            string fname = "Arial";
            if (tsbcbfonts.SelectedIndex != -1)
            {
                fname = tsbcbfonts.SelectedItem.ToString();
            }
            try
            {
                Font font = new Font(new FontFamily(fname), fsize);
                rtbinfo.SelectionFont = font;

            }
            catch
            {
                MessageBox.Show("Font này không hỗ trợ kiểu hiển thị hiện tại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //rtbinfo.Dispose();
        }
        public void readDocFile()
        {
            //VS 2010 ta dùng Application chứ không dùng ApplicationClass.
            // Tạo một thể hiện của ứng dụng MS Word.
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            // Các tham số được sử dụng trong hàm Open được hỗ trợ bởi thư viện API của MS Word
            object fileName = duongdan; //path là đường dẫn đến file cần đọc 
            object missing = System.Reflection.Missing.Value;
            object vk_read_only = false;
            object vk_visible = true;
            object vk_false = false;
            // không sử dụng các thông số không cần thiết ngoai trừ đường dẫn đến file cần mở
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref fileName, ref missing,
                 ref vk_read_only, ref missing,
                 ref missing, ref missing,
                 ref missing, ref missing,
                 ref missing, ref missing,
                 ref missing, ref vk_visible,
                 ref missing, ref missing,
                 ref missing, ref missing);

            doc.ActiveWindow.Selection.WholeStory();
            doc.ActiveWindow.Selection.Copy();
            IDataObject data = Clipboard.GetDataObject();
            rtbinfo.Text = data.GetData(DataFormats.UnicodeText).ToString();//hiển thị dữ liệu lên RichTextBox 

            if (doc != null)
            {
                doc.Close(ref vk_false, ref missing, ref missing);
            }
            wordApp.Quit(ref vk_false, ref missing, ref missing);
        }
        public void taoDinhDangSave()
        {
            save = new SaveFileDialog();
            save.DefaultExt = "rtf";    // Mặc định khi mở của sổ lưu là định dang *.rtf.
            save.Filter = "RichTextFile |*.rtf|Doc file (*.doc)|*.doc|All files (*.*)|*.*";
            //Các định dang khác được cố định khi lưu file.
        }
        public void kiemTraThoat()
        {
            //kiem tra file đã được lưu chưa. nếu = 0, tức là chưa lưu, =1 : tức là đã lưu.
            if (this.kiemTraSave == 0)
            {// nếu chưa lưu. Làm
                if (!rtbinfo.Text.Equals(""))   //kiểm tra nội dung bên trong.
                {//nếu khác rỗng. Làm
                    if (this.duongdan.Equals(""))   //kiểm tra đường dẫn.
                    {//đường dẫn rỗng. Làm
                        save = new SaveFileDialog();
                        save.DefaultExt = "rtf";
                        save.Filter = "RichTextFile |*.rtf";
                        DialogResult result = save.ShowDialog();
                        /*
                         * mặc định khi show ra cửa sổ OpenFileDialog và SaveFileDialog.
                         * sẽ có 3 button OK / NO / Cancel.
                         */
                        if (result == DialogResult.Cancel)
                        {//nếu là Cancel. Làmd
                            return;     //trả về và không làm gì. thoát khỏi sự kiện.
                        }
                        //gán đường dẫn.
                        this.duongdan = save.FileName;
                        try
                        {
                            rtbinfo.SaveFile(duongdan); //lưu file theo đường dẫn đã chọn
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString(), "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                    }
                    else     //Ngược lại trường hợp đường dẫn rỗng.
                    {
                        try
                        {
                            //có đường dẫn rồi. cho phép lưu chồng file theo đường dẫn.
                            rtbinfo.SaveFile(duongdan);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString(), "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        }
                    }
                }
            }
        }
        public void kiemTraNew()
        {
            //xuất hiện thông báo với Icon Question
            if (!rtbinfo.Text.Equals(""))
            {
                DialogResult chon = MessageBox.Show("Bạn muốn lưu trước khi tạo mới !", "Thông Báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                // Nếu chọn ok
                if (chon == DialogResult.Yes)
                {
                    /* * Làm.
                     *  Nếu đường dẫn rỗng. trong lập tình hướng đối tượng không thể so sánh chuỗi bằng '=='. ta phải dùng từ khóa 'Equals'
                    * */
                    if (this.duongdan.Equals(""))
                    {
                        // Khởi tạo sự kiện lưu file.
                        save = new SaveFileDialog();
                        save.DefaultExt = "rtf";
                        save.Filter = "RichTextFile |*.rtf";    //Gán cho file lưu xuống mặc định là *.rtf.
                        /* *
                         * Gọi chức năng lưu của hệ thống Window
                         * */
                        DialogResult result = save.ShowDialog();
                        // Nếu chọn cancel.
                        if (result == DialogResult.Cancel)
                        {
                            //kiểu trả về, rời khỏi sự kiện thoát.
                            return;
                        }
                        /** Nếu không chọn cancel.
                         *  Gán duongdan = với save.Filename;
                        */
                        this.duongdan = save.FileName;

                        // Lưu đường dẫn vào biến.
                        try     // Thử.
                        {
                            rtbinfo.SaveFile(duongdan);     // Lưu tập tin từ RichTextBox định dạng mặc định [*.rtf]
                        }
                        catch (System.Exception ex)     // Trường hợp ngoại lệ
                        {
                            MessageBox.Show(ex.Message.ToString(), "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                            // Thông báo lỗi và rời khỏi sự kiện.
                        }
                    }
                    else
                    {
                        //Tức là có đường dẫn không trống hay là đã có đường dẫn.
                        // Tiến hành lưu tập tin mà không cần gọi cửa sổ đê lưu
                        try
                        {
                            rtbinfo.SaveFile(duongdan);     // Lưu tập tin.
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString(), "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        }
                    }
                    rtbinfo.Text = "";     // Khởi tạo RichTextBox.
                }
                // Ngược lại trường hợp chọn Yes
                else if (chon == DialogResult.No)
                {
                    rtbinfo.Text = "";     // Thoát không cần lưu.
                }
                else    // Ngược lại nếu chọn nút thứ 3 là nút Cancel.
                {
                    return;     // Trả về, không lưu gì, thoát khỏi sự kiện.
                }

            }
            else
            {
                return;
            }
        }
        public void loadFonts()
        {
            //tạo danh sách cỡ chữ....
            for (int i = 8; i <= 72; i++)
            {
                tsbcbfontsize.Items.Add(i.ToString());
            }
            //set mặc định cho Size.
            tsbcbfontsize.SelectedIndex = 4;
            System.Drawing.Text.InstalledFontCollection fonts = new System.Drawing.Text.InstalledFontCollection();
            //tao danh sach Fonts lấy từ hệ thống chính máy tính bạn.
            foreach (FontFamily f in fonts.Families)
            {
                tsbcbfonts.Items.Add(f.Name.ToString());
            }
            tsbcbfonts.SelectedItem = "Arial";
            //set mặc định cho Fonts.
        }
        public void kiemTraOpen()
        {
            open = new OpenFileDialog();
            save = new SaveFileDialog();
            if (!this.rtbinfo.Text.Equals(""))
            {
                DialogResult Result2 = MessageBox.Show("Bạn có muốn lưu lại trước khi mở tệp mới ?", "Thông báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                //tạo một biến kết quả trả về từ MessageBox
                if (Result2 == DialogResult.Yes) //nếu chọn Yes
                {
                    //tương tự menu item Thoát
                    if (this.duongdan.Equals(""))
                    {
                        DialogResult resVal = save.ShowDialog();
                        if (resVal == DialogResult.Cancel)
                        {
                            return;
                        }
                        this.duongdan = save.FileName;

                        try
                        {
                            rtbinfo.SaveFile(duongdan);
                        }
                        catch (Exception Ex)
                        {
                            MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                    }
                    else
                    {
                        try
                        {
                            rtbinfo.SaveFile(duongdan);
                        }
                        catch (Exception Ex)
                        {
                            MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        }
                    }

                    DialogResult Result = open.ShowDialog();
                    //tạo một biến kết quả từ việc mở tập tin
                    //lúc đó hộp thoại mở tập tin sẽ có 2 nút là Open và Cancel
                    if (Result == DialogResult.Cancel) //nếu chọn Cancel
                    {
                        return; //rời khỏi sự kiện
                    }
                    else //hoặc chọn mở
                    {
                        try //thử
                        {
                            this.duongdan = open.FileName; //lưu đường dẫn tập tin
                            readDocFile();
                        }
                        catch (Exception Ex) //nếu có ngoại lệ
                        {
                            MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            //thông báo lổi
                        }
                    }
                }
                else if (Result2 == DialogResult.No)
                {
                    DialogResult Result = open.ShowDialog();
                    //tạo một biến kết quả từ việc mở tập tin
                    //lúc đó hộp thoại mở tập tin sẽ có 2 nút là Open và Cancel
                    if (Result == DialogResult.Cancel) //nếu chọn Cancel
                    {
                        return; //rời khỏi sự kiện
                    }
                    else //hoặc chọn mở
                    {
                        try //thử
                        {
                            this.duongdan = open.FileName; //lưu đường dẫn tập tin
                            readDocFile();
                        }
                        catch (Exception Ex) //nếu có ngoại lệ
                        {
                            MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            //thông báo lổi
                        }
                    }
                }
                else
                {
                    return; //rời khỏi sự kiện
                }
            }
            else
            {
                DialogResult Result = open.ShowDialog();
                //tạo một biến kết quả từ việc mở tập tin
                //lúc đó hộp thoại mở tập tin sẽ có 2 nút là Open và Cancel
                if (Result == DialogResult.Cancel) //nếu chọn Cancel
                {
                    return; //rời khỏi sự kiện
                }
                else //hoặc chọn mở
                {
                    try //thử
                    {
                        this.duongdan = open.FileName; //lưu đường dẫn tập tin
                        readDocFile();
                    }
                    catch (Exception Ex) //nếu có ngoại lệ
                    {
                        MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        //thông báo lổi
                    }
                }
            }
        }
        public void saveData()
        {

            //giống như trên nhưng không cần biết đã lưu tập tin hay chưa
            if (this.duongdan.Equals(""))
            {
                save = new SaveFileDialog();
                save.DefaultExt = "rtf";
                save.Filter = "RichTextFile |*.rtf";
                DialogResult Result = save.ShowDialog();

                if (Result == DialogResult.Cancel)
                {
                    return;
                }

                duongdan = save.FileName;

                try
                {
                    rtbinfo.SaveFile(duongdan);
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(Ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }


            this.kiemTraSave = 1; //đặt lại trạng thái cho biến kiểm tra việc lưu

        }
        #endregion

        #region Bắt Các Sự Kiện CLick Trong MenuStrip
        //MenuStrip Exit.
        private void menustrip_Exit_Click(object sender, EventArgs e)
        {
            //kiem tra 
            if (this.kiemTraSave == 0)   //Tức là đoạn văn bản chưa được lưu.
            {
                Application.Exit();
                /************************************************************************/
                /** Khi thoát chương trình.
                 *  Mặc định hệ thống sẽ gọi sự kiện FormClosing hoặc FormClosed.
                 *  Để tránh trường hợp này, ta nên lập trình cho sự kiện thoát của FORM trước.
                 *  Sau đó vào phần code cho Menu chỉ cần khọi hàm thoát, khi đó chương trình sẽ tự động gọi hàm.
                 *  FormClosing hoặc FormClosed.
                 */
                /************************************************************************/
            }
            // Ngược lại trường hợp kiemTraSave = 0
            // tức là đã lưu rồi.
            else
            {
                Application.Exit();
            }
        }
        //MenuStrip Save As
        private void menustrip_SaveAs_Click(object sender, EventArgs e)
        {
            taoDinhDangSave();
            DialogResult chon = save.ShowDialog();
            this.duongdan = save.FileName;
            try
            {
                if (chon == DialogResult.OK)
                {
                    rtbinfo.SaveFile(duongdan);
                }
                else
                {
                    return;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //MenuStrip Save.
        private void menustrip_Save_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                saveData();
            }
            else
            {
                rtbinfo.SaveFile(duongdan);
            }
        }
        //MenuStrip Open.
        private void menustrip_Open_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                kiemTraOpen();
            }
        }
        //MenuStrip New.
        private void menustrip_New_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                kiemTraNew();
            }
            else
            {
                rtbinfo.Text = "";  // Tạo mới
            }
        }
        //MenuStrip Copy.
        private void menustrip_Copy_Click(object sender, EventArgs e)
        {
            rtbinfo.Copy();
        }
        //MenuStrip Cut.
        private void menustrip_Cut_Click(object sender, EventArgs e)
        {
            rtbinfo.Cut();
        }
        //MenuStrip Undo.
        private void menustrip_Undo_Click(object sender, EventArgs e)
        {
            rtbinfo.Undo();
        }
        //MenuStrip Redo.
        private void menustrip_Re_Undo_Click(object sender, EventArgs e)
        {
            rtbinfo.Redo();
        }
        //MenuStrip In Đậm.
        private void menustrip_InDam_Click(object sender, EventArgs e)
        {
            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Bold);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }
        }
        //MenuStrip In Nghiêng.
        private void menustrip_InNghieng_Click(object sender, EventArgs e)
        {
            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Italic);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }
        }
        //MenuStrip Gạch Chân.
        private void menustrip_Gach_Chan_Click(object sender, EventArgs e)
        {
            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Underline);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }
        }
        //MenuStrip Canh Giữa.
        private void menustrip_Canh_Giua_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Center;
        }
        //MenuStrip Canh Trái.
        private void menustrip_Canh_Trai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Left;
        }
        //MenuStrip Canh Phải.
        private void menustrip_Canh_Phai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Right;
        }
        //MenuStrip Chèn Hình.
        private void menustrip_Chen_Hinh_Click_1(object sender, EventArgs e)
        {
            open = new OpenFileDialog();
            DialogResult Result = open.ShowDialog();
            //khai báo biến kết quả cho việc mở tập tin
            if (Result == DialogResult.Cancel) //nếu chọn Cancel
            {
                return; //rời khỏi sự kiện
            }
            else //hoặc chọn Open
            {
                try //thử
                {
                    string ImagePath = open.FileName; //lấy đường dẫn 
                    //của tập tin

                    Bitmap myBitmap = new Bitmap(ImagePath); //tạo một Bitmap

                    Clipboard.SetDataObject(myBitmap); //đặt đối tượng dử liệu vào Clipboard

                    DataFormats.Format myFormat = DataFormats.GetFormat(DataFormats.Bitmap);
                    //lấy định dạng của hình
                    if (rtbinfo.CanPaste(myFormat)) //nếu có thể chèn vào RTBox
                    {
                        rtbinfo.Paste(myFormat); //chèn hình vào
                    }
                    else //nếu không thể chèn
                    {
                        MessageBox.Show("Không thể chèn !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        //hiển thị thông báo
                    }
                }
                catch (Exception ex) //nếu có ngoại lệ (mở tập tin không phải hình)
                {
                    MessageBox.Show(ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                    //thông báo và rời khỏi sự kiện
                }
            }
        }
        //MenuStrip Đánh Dấu.
        private void menustrip_Danh_Dau_Click(object sender, EventArgs e)
        {
            if (check == 0)
            {
                rtbinfo.SelectionBullet = true; //đánh dấu
                check++;
            }
            else
            {
                rtbinfo.SelectionBullet = false; //hủy đánh dấu
                check--;
            }
        }
        //MenuStrip Chọn Màu Chữ.
        private void menustrip_Mau_Chu_Click(object sender, EventArgs e)
        {
            color = new ColorDialog();
            color.ShowDialog(); //hiển thị hộp thoại chọn màu
            rtbinfo.SelectionColor = color.Color; //đặt lại màu
        }
        //MenuStrip Hướng Dẫn
        private void menustrip_huongdan_Click(object sender, EventArgs e)
        {
            frmhuongdan huongdan = new frmhuongdan();
            huongdan.Show();
        }
        //MenuStrip Thông Tin
        private void menustrip_thongtin_Click(object sender, EventArgs e)
        {
            ThongTin thongtin = new ThongTin();
            thongtin.Show();
        }
        //MenuStrip Search.
        private void menustrip_Search_Click(object sender, EventArgs e)
        {
            if (fsearch == null || fsearch.IsDisposed)
                fsearch = new frmtimkiem(rtbinfo);
            fsearch.ShowFind();
        }
        #endregion
        
        #region Bắt Sự Kiện CLick ToolStrip
        //ToolStrip Combobox Chọn Fonts.
        private void tsbcbfonts_Click(object sender, EventArgs e)
        {
            dinhDangFonts();
        }
        //ToolStrip Combobox Chọn Size.
        private void tsbcbfontsize_Click(object sender, EventArgs e)
        {
            dinhDangFonts();
        }
        //ToolStrip Icon New (tạo file mới trong RichTextBox)
        private void tsbnew_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                kiemTraNew();
            }
            else
            {
                rtbinfo.Text = "";  // Tạo mới
            }
        }
        //ToolStrip Icon Save.
        private void tsbsave_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                saveData();
            }
            else
            {
                rtbinfo.SaveFile(duongdan);
            }
        }
        //ToolStrip Icon Open.
        private void tsbopen_Click(object sender, EventArgs e)
        {
            if (kiemTraSave == 0)
            {
                kiemTraOpen();
            }
        }
        //ToolStrip Icon Undo
        private void tsbundo_Click(object sender, EventArgs e)
        {
            rtbinfo.Undo();
        }
        //ToolStrip Icon Re Undo
        private void tsbreundo_Click(object sender, EventArgs e)
        {
            rtbinfo.Redo();
        }
        //ToolStrip Icon Cut
        private void tsbcut_Click(object sender, EventArgs e)
        {
            rtbinfo.Cut();
        }
        //ToolStrip Icon Copy
        private void tsbcopy_Click(object sender, EventArgs e)
        {
            rtbinfo.Copy();
        }
        //ToolStrip Icon Paste
        private void tsbpaste_Click(object sender, EventArgs e)
        {
            rtbinfo.Paste();
        }
        //ToolStrip Icon In Đậm.
        private void tsbindam_Click(object sender, EventArgs e)
        {

            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Bold);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }

        }
        //ToolStrip Icon In Nghiêng.
        private void tsbinnghieng_Click(object sender, EventArgs e)
        {
            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Italic);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }
        }
        //ToolStrip Icon Gạch Chân.
        private void tsbgachchan_Click(object sender, EventArgs e)
        {
            if (this.check == 0)
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Underline);
                check++;
            }
            else
            {
                rtbinfo.SelectionFont = new Font(rtbinfo.SelectionFont, FontStyle.Regular);
                check--;
            }
        }
        //ToolStrip Icon Canh Trái.
        private void tsbcanhtrai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Left;
        }
        //ToolStrip Icon Canh Giữa.
        private void tsbcanhgiua_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Center;
        }
        //ToolStrip Icon Canh Phải.
        private void tsbcanhphai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Right;
        }
        //ToolStrip Icon Chèn Hình.
        private void tsbchonhinh_Click(object sender, EventArgs e)
        {
            open = new OpenFileDialog();
            DialogResult Result = open.ShowDialog();
            //khai báo biến kết quả cho việc mở tập tin
            if (Result == DialogResult.Cancel) //nếu chọn Cancel
            {
                return; //rời khỏi sự kiện
            }
            else //hoặc chọn Open
            {
                try //thử
                {
                    string ImagePath = open.FileName; //lấy đường dẫn 
                    //của tập tin

                    Bitmap myBitmap = new Bitmap(ImagePath); //tạo một Bitmap

                    Clipboard.SetDataObject(myBitmap); //đặt đối tượng dử liệu vào Clipboard

                    DataFormats.Format myFormat = DataFormats.GetFormat(DataFormats.Bitmap);
                    //lấy định dạng của hình
                    if (rtbinfo.CanPaste(myFormat)) //nếu có thể chèn vào RTBox
                    {
                        rtbinfo.Paste(myFormat); //chèn hình vào
                    }
                    else //nếu không thể chèn
                    {
                        MessageBox.Show("Không thể chèn !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        //hiển thị thông báo
                    }
                }
                catch (Exception ex) //nếu có ngoại lệ (mở tập tin không phải hình)
                {
                    MessageBox.Show(ex.Message.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                    //thông báo và rời khỏi sự kiện
                }
            }
        }
        //ToolStrip Icon Đánh Dấu.
        private void tsbdanhdau_Click(object sender, EventArgs e)
        {
            if (check == 0)
            {
                rtbinfo.SelectionBullet = true; //đánh dấu
                check++;
            }
            else
            {
                rtbinfo.SelectionBullet = false; //hủy đánh dấu
                check--;
            }

        }
        //ToolStrip Icon Màu Chữ.
        private void tsbchonmauchu_Click(object sender, EventArgs e)
        {
            color = new ColorDialog();
            DialogResult mau = color.ShowDialog(); //hiển thị hộp thoại chọn màu
            if (mau == DialogResult.OK)
            {
                rtbinfo.SelectionColor = color.Color; //đặt lại màu
                toolStrip_mau.BackColor = color.Color;
            }
            else
            {
                return;
            }
            
        }
        //ToolStrip Icon Fonts Full.
        private void tsbfontstyle_Click(object sender, EventArgs e)
        {
            dialogFonts = new FontDialog();
            DialogResult font = dialogFonts.ShowDialog();
            if (this.rtbinfo.SelectedText.Equals(""))
            {
                if (font == DialogResult.OK)
                {
                    rtbinfo.Font = dialogFonts.Font;
                }
                else
                {
                    return;
                }

            }
            else
            {
                if (font == DialogResult.OK)
                {
                    rtbinfo.SelectionFont = dialogFonts.Font;
                }
                else
                {
                    return;
                }
            }
        }
        //ToolStrip Icon Help.
        private void tsbhelp_Click(object sender, EventArgs e)
        {
            frmhuongdan huongdan = new frmhuongdan();
            huongdan.Show();

        }
        //ToolStrip Icon Thông Tin.
        private void tsbthongtin_Click(object sender, EventArgs e)
        {
            ThongTin info = new ThongTin();
            info.Show();
        }
        #endregion

       
        private void tsbtimkiem_Click(object sender, EventArgs e)
        {
            if (fsearch == null || fsearch.IsDisposed)
                fsearch = new frmtimkiem(rtbinfo);
            fsearch.ShowFind();
        }

        #region Control Right CLick Mouse
        private void rightclick_cut_Click(object sender, EventArgs e)
        {
            rtbinfo.Cut();
        }

        private void rightclick_copy_Click(object sender, EventArgs e)
        {
            rtbinfo.Copy();
        }

        private void rightclick_paste_Click(object sender, EventArgs e)
        {
            rtbinfo.Paste();
        }

        private void rightclick_fonts_Click(object sender, EventArgs e)
        {
            dialogFonts = new FontDialog();
            DialogResult font = dialogFonts.ShowDialog();
            if (this.rtbinfo.SelectedText.Equals(""))
            {
                if (font == DialogResult.OK)
                {
                    rtbinfo.Font = dialogFonts.Font;
                }
                else
                {
                    return;
                }

            }
            else
            {
                if (font == DialogResult.OK)
                {
                    rtbinfo.SelectionFont = dialogFonts.Font;
                }
                else
                {
                    return;
                }
            }
        }

        private void rightclick_danhdau_Click(object sender, EventArgs e)
        {
            if (check == 0)
            {
                rtbinfo.SelectionBullet = true; //đánh dấu
                check++;
            }
            else
            {
                rtbinfo.SelectionBullet = false; //hủy đánh dấu
                check--;
            }
        }

        private void rightclick_mauchu_Click(object sender, EventArgs e)
        {
            color = new ColorDialog();
            color.ShowDialog(); //hiển thị hộp thoại chọn màu
            rtbinfo.SelectionColor = color.Color; //đặt lại màu
        }

        private void rightclick_canhtrai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void rightclick_canhgiua_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void rightclick_canhphai_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void rightclick_selectall_Click(object sender, EventArgs e)
        {
            rtbinfo.SelectAll();

        }
        #endregion

        private void frmworkpad_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult thoat = MessageBox.Show("Bạn muốn thoát khỏi chương trinh", "Thông Báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (thoat == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
            else
            {
                kiemTraThoat();
            }

        }
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
