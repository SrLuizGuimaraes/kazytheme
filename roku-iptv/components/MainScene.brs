' Playlist M3U carregada na abertura. Troque aqui para usar outra lista.
function playlistUrl() as string
    return "https://iptv-org.github.io/iptv/countries/br.m3u"
end function

sub init()
    m.list = m.top.findNode("list")
    m.video = m.top.findNode("video")
    m.status = m.top.findNode("status")

    m.list.observeField("itemSelected", "onChannelSelected")
    m.video.observeField("state", "onVideoState")

    m.task = CreateObject("roSGNode", "PlaylistTask")
    m.task.url = playlistUrl()
    m.task.observeField("content", "onPlaylistLoaded")
    m.task.observeField("error", "onPlaylistError")
    m.task.control = "RUN"
end sub

sub onPlaylistLoaded()
    m.channels = m.task.content
    if m.channels = invalid or m.channels.getChildCount() = 0 then
        m.status.text = "Nenhum canal encontrado na lista."
        return
    end if
    m.list.content = m.channels
    m.status.visible = false
    m.list.visible = true
    m.list.setFocus(true)
end sub

sub onPlaylistError()
    m.status.text = "Erro ao carregar a lista: " + m.task.error
end sub

sub onChannelSelected()
    item = m.channels.getChild(m.list.itemSelected)
    if item = invalid then return
    m.video.content = item
    m.video.visible = true
    m.video.setFocus(true)
    m.video.control = "play"
end sub

sub onVideoState()
    if m.video.state = "error" then
        stopVideo()
        m.status.text = "Canal indisponível. Escolha outro."
        m.status.visible = true
    end if
end sub

sub stopVideo()
    m.video.control = "stop"
    m.video.visible = false
    m.list.setFocus(true)
end sub

function onKeyEvent(key as string, press as boolean) as boolean
    if press and key = "back" and m.video.visible then
        m.status.visible = false
        stopVideo()
        return true
    end if
    return false
end function
