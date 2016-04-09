
Module Mod_Audio

    'play audio file ------------------------
    'Api to send the commands to the mci device
    Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long 'Get the error message of the mcidevice if any
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long 'Send command strings to the mci device

    'play audio file-----------------------
    Private Function openMovie(ByVal Filename$) As Long

        Dim _Error&

        'Open a movie in the default window style(Popup)
        Filename = Chr(34) & Filename & Chr(34)
        _Error = mciSendString("close movie", 0, 0, 0)
        'Decide which way you want the mci device to work below

        'Specify the mpegvideo driver to play the movies
        _Error = mciSendString("open " & Filename & " type mpegvideo alias movie", 0, 0, 0)

        'Let the mci device decide which driver to use
        'Error = mciSendString("open " & Filename & " alias movie", 0, 0, 0)

        openMovie = _Error
    End Function

    Private Function playMovie() As Long
        Dim _Error&
        'Play the movie after you open it
        _Error = mciSendString("play movie", 0, 0, 0)
        playMovie = _Error
    End Function

    Private Function stopMovie() As Long
        Dim _Error&
        'Stops the playing of the movie
        _Error = mciSendString("stop movie", 0, 0, 0)
        stopMovie = _Error
    End Function

    Private Function closeMovie() As Long
        Dim _Error&
        'Close the mci device
        _Error = mciSendString("close all", 0, 0, 0)
        closeMovie = _Error
    End Function



    Public Sub PlayAudio(ByVal Filename$) ' Play
        stopMovie()
        closeMovie()
        openMovie(Filename)
        playMovie()
    End Sub

    Public Sub StopAudio() ' Stop
        stopMovie()
        closeMovie()
    End Sub

    Public Function setVolume(ByVal Value As Long) As Long
        Dim _Error As Long
        'Raise or lower the volume for both channels
        '1000 max - 0 min
        _Error = mciSendString("setaudio movie volume to " & Value, 0, 0, 0)
        setVolume = _Error
    End Function


End Module
