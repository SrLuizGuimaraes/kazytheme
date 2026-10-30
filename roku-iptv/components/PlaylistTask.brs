sub init()
    m.functionName = "fetchPlaylist"
end sub

sub fetchPlaylist()
    req = CreateObject("roUrlTransfer")
    req.setUrl(m.top.url)
    req.setCertificatesFile("common:/certs/ca-bundle.crt")
    req.initClientCertificates()
    req.enableEncodings(true)
    req.retainBodyOnError(true)

    body = req.getToString()
    if body = invalid or Len(body) = 0 then
        m.top.error = "sem resposta do servidor"
        return
    end if
    if Left(body, 7) <> "#EXTM3U" then
        m.top.error = "arquivo nao parece uma playlist M3U"
        return
    end if

    m.top.content = parseM3U(body)
end sub

function parseM3U(body as string) as object
    root = CreateObject("roSGNode", "ContentNode")
    title = ""

    for each rawLine in body.split(chr(10))
        line = rawLine.trim()
        if Left(line, 8) = "#EXTINF:" then
            comma = line.instr(",")
            if comma >= 0 then title = line.mid(comma + 1).trim() else title = "Canal"
        else if line <> "" and Left(line, 1) <> "#" then
            item = CreateObject("roSGNode", "ContentNode")
            item.title = title
            item.url = line
            item.live = true
            if line.instr(".m3u8") >= 0 then item.streamFormat = "hls"
            root.appendChild(item)
            title = ""
        end if
    end for

    return root
end function
