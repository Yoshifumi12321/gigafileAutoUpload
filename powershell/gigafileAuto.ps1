# メイン処理
function Main($args)
{
    # URLを入力する
    $url = "https://gigafile.nu/"

    # シェル
    $shell = New-Object -ComObject Shell.Application
    # IEを起動する
    $ie = New-Object -ComObject InternetExplorer.Application

    # HttpUtilityクラスの有効化
    $encode = Add-Type -AssemblyName System.Web
    #
    # $encode = [System.Text.Encoding]::default

    # HWND
    $hwnd = $ie.HWND

    if ($ie -ne $null)
    {
        # IEを表示する
        $ie.Visible = $true

        # URLを開く(キャッシュ無効)
        $ie.Navigate($url, 4)

        # Webページの読み込みが終わるまで待機する
        Wait-LoadedPage $ie

        # HTMLドキュメントを取得する
        $doc = $ie.Document

        if ($doc -ne $null)
        {
            # # テキストボックスをIDで探し、値をセットする
            # $file_lifetime =$doc.getElementsById("file_lifetime")
            # if ($file_lifetime -ne $null)
            # {
            #     $file_lifetime.value = 60
            # }

            # テキストボックスをIDで探し、値をセットする
            $file_lifetime =[System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @("file_lifetime"))
            if ($file_lifetime -ne $null)
            {
                $file_lifetime.value = 60
            }

            $lifetime_meter =[System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @("lifetime_meter"))
            if ($lifetime_meter -ne $null)
            {
                $lifetime_meter.style.margin = "inherit inherit inherit 204px"
                $lifetime_meter.style.width = "100%"
                $lifetime_meter.style.textAlign = "center"
            }

            $lifetime_text =[System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @("lifetime_text"))
            if ($lifetime_text -ne $null)
            {
                $lifetext = [System.Web.HttpUtility]::UrlEncode("60 days")
                $lifetime_text.innerHTML = "60 days"
            }

            # ファイル選択ボタンをNameで探し、押下する
            $select_btn = [System.__ComObject].InvokeMember("getElementById",[System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @("select_btn"))
            if ($select_btn -ne $null)
            {
                $select_btn.click()
            }

            $file_url = [System.__ComObject].InvokeMember("getElementByClassName",[System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @("file_info_url"))
            if ($file_url -ne $null)
            {
                echo $file_url.value > a.txt
            }

            # Webページの読み込みが終わるまで待機する
            Wait-LoadedPage $ie
        }
        #
        # # IEを終了する
        # $ie.Quit()
        #
        # # IEを破棄する
        # $ie = $null
    }
}

# Webページの読み込み完了を待つ
function Wait-LoadedPage($ie)
{
    # Webページの読み込みが終わるまで待機する
    while ($ie.busy -or $ie.readystate -ne 4)
    {
        Start-Sleep -Milliseconds 100
    }
}

# エントリーポイント
Main $args
