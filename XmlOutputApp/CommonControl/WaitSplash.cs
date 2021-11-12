using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WorldGyomu.CommonControl
{
    public partial class WaitSplash : Form
    {
        // スプラッシュ表示用フォーム
        private static WaitSplash _splashForm = null;
        // スプラッシュ表示用スレッド
        private static System.Threading.Thread _thread = null;
        // ロック用オブジェクト
        private static readonly object syncObject = new object();
        // スプラッシュ表示の待機用ハンドル
        private static System.Threading.ManualResetEvent splashShownEvent = null;

        public WaitSplash()
        {
            InitializeComponent();
        }

        public static WaitSplash Form
        {
            get { return _splashForm; }
        }

        // スプラッシュフォーム表示(別スレッドで表示)
        public static void ShowSplash()
        {
            lock (syncObject)
            {
                if (_splashForm == null || _thread == null)
                {
                    //Application.IdleイベントハンドラでSplashフォームを閉じる
                    Application.Idle += new EventHandler(Application_Idle);

                    //待機ハンドルの作成
                    splashShownEvent = new System.Threading.ManualResetEvent(false);

                    //スレッドの作成
                    _thread = new System.Threading.Thread(
                        new System.Threading.ThreadStart(StartThread));
                    _thread.Name = "SplashForm";
                    _thread.IsBackground = true;    // バックグラウンドスレッド
                    _thread.SetApartmentState(System.Threading.ApartmentState.STA);
                    //スレッドの開始
                    _thread.Start();
                }
            }
        }

        public static void CloseSplash()
        {
            lock (syncObject)
            {
                if (_thread == null)
                {
                    return;
                }

                //Application.Idleイベントハンドラの削除
                Application.Idle -= new EventHandler(Application_Idle);

                //Splashが表示されるまで待機する
                if (splashShownEvent != null)
                {
                    splashShownEvent.WaitOne();
                    splashShownEvent.Close();
                    splashShownEvent = null;
                }

                //Splashフォームを閉じる
                //Invokeが必要か調べる
                if (_splashForm != null)
                {
                    if (_splashForm.InvokeRequired)
                    {
                        _splashForm.Invoke(new MethodInvoker(CloseSplashForm));
                    }
                    else
                    {
                        CloseSplashForm();
                    }
                    _splashForm = null;
                }
            }
        }

        public static void CloseSplashForm()
        {
            //Splashフォームがあるか調べる
            if (_splashForm.IsDisposed == false)
            {
                //Splashフォームを閉じる
                _splashForm.Close();
            }
        }

        // アプリケーションがアイドル状態になった場合の処理
        private static void Application_Idle(object sender, EventArgs e)
        {
            CloseSplash();
        }

        // スレッドの処理開始
        private static void StartThread()
        {
            //Splashフォームを作成
            _splashForm = new WaitSplash();
            //Splashが表示されるまでCloseSplashメソッドをブロックする
            _splashForm.Activated += new EventHandler(_form_Activated);
            //Splashフォームを表示する
            Application.Run(_splashForm);
        }

        // スプラッシュのフォームが表示された場合の処理
        private static void _form_Activated(object sender, EventArgs e)
        {
            _splashForm.Activated -= new EventHandler(_form_Activated);
            //CloseSplashメソッドの待機を解除
            if (splashShownEvent != null)
            {
                splashShownEvent.Set();
            }
        }

		private void attention_Click(object sender, EventArgs e)
		{
        
		}
	}
}
