using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Chuong_Trinh_Soan_Thao_Van_Ban
{
    public partial class frmtimkiem : Form
    {
        /************************************************************************/
        /* lưu ý trước khi tạo class này ta cần thức hiện thao tác click vào RichTextBox bên Form Chính.
         * cài đặt lại thuộc tính RichTextBox là Hideselection = false.
         * để khi bấm vào nút tìm kiếm bên class này, thì class Form chính trong RichTex sẽ không bị ẩn đi vùng tìm kiếm.
        /************************************************************************/

        private RichTextBox _richTextBox;            //khởi tạo biến bool _close = 'True' or 'false' để kiểm tra.
        // Khởi tạo có tham số. Để truyền RichTextBox là rtbinfo trong class chương trình demo.
        public frmtimkiem(RichTextBox richTextBox)
        {
            InitializeComponent();
            _richTextBox = richTextBox;
        }
        public void ShowFind()
        {
            //Hiển thị Form Search với tham số là RichTextBox được truyền vào.
            this.Show(_richTextBox);
            txtsearch.Focus();
            txtsearch.SelectAll();
        }
        // Bắt sự kiện Click cho button Search.
        public void btnsearch_Click(object sender, EventArgs e)
        {
            //RichTextBox hỗ trợ phương Thức int Find(....);
            Find(_richTextBox, txtsearch.Text, radup.Checked);
        }



        public void Find(RichTextBox richtext, string nhap, bool check_up)
        {
            //Khởi tạo giá trị RichTextBoxFinds = none.
            RichTextBoxFinds chose = RichTextBoxFinds.None;
            //nếu đã check vào radup. hay raduo.checked
            if (check_up)
            {
                //tìm kiếm chạy ngược từ dưới lên trên
                chose |= RichTextBoxFinds.Reverse;
            }
            // tạo biến int. như đã nói trên RichTextBox hỗ trợ phương thức int Find(...). giá trị trả về là int.
            int index;
            //neu check_up = true
            if (check_up)
            {
                index = richtext.Find(nhap,0, richtext.SelectionStart, chose);
            }
            else
            {
                index = richtext.Find(nhap, richtext.SelectionStart + richtext.SelectionLength, chose);
            }
            if (index >= 0)
            {
                richtext.SelectionStart = index;
                richtext.SelectionLength = nhap.Length;
            }
            else
            {
                MessageBox.Show(Application.ProductName + " has finished searching the document.",
                                Application.ProductName, MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }
        }
    }
}
