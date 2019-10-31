using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mntsiwake
{
    class errEnd
    {
        //エラー後終了時の処理
        public static void Exit()
        {
            utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.TMPREAD);
            utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.tmpFile);
    
            MessageBox.Show ("変換処理は中断されました。",Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Information);
            Environment.Exit(0);
        }
    }
}
