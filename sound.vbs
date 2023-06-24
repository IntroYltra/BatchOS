Set WMP = WScript.CreateObject("MediaPlayer.MediaPlayer","WMP_")
WMP.Open "sound.wav"
WMP.Play
WScript.Sleep 5000