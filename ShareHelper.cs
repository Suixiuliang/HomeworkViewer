using System.Threading.Tasks;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public static class ShareHelper
    {
        public static async Task ShowShareUIAsync(string filePath)
        {
            // 分享功能需要 Windows 11 且项目配置正确，这里仅提示
            MessageBox.Show($"文件已保存至：{filePath}\n分享功能暂不可用，请手动分享。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            await Task.CompletedTask;
        }
    }
}