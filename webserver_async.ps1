# New-ScriptblockCallbackは、Register-ObjectEventを使った非同期処理の要のコード
# Polaris　から移植
# https://powershell.github.io/Polaris/docs/api/New-ScriptblockCallback.html
# https://github.com/PowerShell/Polaris.git
# <---  New-ScriptblockCallback.ps1 から必要部を移植　開始    --->
function New-ScriptblockCallback {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [scriptblock]$Callback
    )

    # is this type already defined?
    if (-not ("CallbackEventBridge" -as [type])) {
        Add-Type @"
            using System;
            public sealed class CallbackEventBridge
            {
                public event AsyncCallback CallbackComplete = delegate { };
                private CallbackEventBridge() {}
                private void CallbackInternal(IAsyncResult result)
                {
                    CallbackComplete(result);
                }
                public AsyncCallback Callback
                {
                    get { return new AsyncCallback(CallbackInternal); }
                }
                public static CallbackEventBridge Create()
                {
                    return new CallbackEventBridge();
                }
            }
"@
    }
    $bridge = [callbackeventbridge]::create()
    Register-ObjectEvent -input $bridge -EventName callbackcomplete -action $callback -messagedata $args > $null
    $bridge.callback
}
# <---  New-ScriptblockCallback.ps1 から必要部を移植　終了    --->
# もし、New-ScriptblockCallback.ps１をリンクする場合は、移植コードをカットし、↓をコメントアウト
#. ".\\New-ScriptblockCallback.ps1"


$listener = New-Object Net.HttpListener
$listener.Prefixes.Add("http://localhost:8000/")
#$listener.Prefixes.Add("http://+:80/Temporary_Listen_Addresses/")

# 対応する非同期操作が完了したときに呼び出されるコールバックメソッド
$ListenerCallback = (New-ScriptblockCallback -Callback {
            param(
                [System.IAsyncResult]
                $AsyncResult
            )

            [Net.HttpListener]$listener = $AsyncResult.AsyncState
            $Context = $Listener.EndGetContext($AsyncResult)
            $response = $context.Response
            $content = [System.Text.Encoding]::UTF8.GetBytes('hello world! by ListenerCallback!')
            $response.OutputStream.Write($content, 0, $content.Length)
            $response.Close()

        })

try {

    $listener.Start()
    $IAsyncResult = $Listener.BeginGetContext($ListenerCallback, $Listener)
    write-host 'Waiting for request to be processed asyncronously.'
    $IAsyncResult.AsyncWaitHandle.WaitOne()
    write-host 'Request processed asyncronously.'
    $listener.Close();
    
}
catch {
    Write-Error($_.Exception)
}