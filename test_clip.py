import win32clipboard

def copy_html_to_clipboard(html: str):
    header = (
        "Version:0.9\r\n"
        "StartHTML:{0:08d}\r\n"
        "EndHTML:{1:08d}\r\n"
        "StartFragment:{2:08d}\r\n"
        "EndFragment:{3:08d}\r\n"
    )
    prefix = "<html><body><!--StartFragment-->"
    suffix = "<!--EndFragment--></body></html>"
    
    html_bytes = html.encode('utf-8')
    prefix_bytes = prefix.encode('utf-8')
    suffix_bytes = suffix.encode('utf-8')
    
    start_html = 105
    start_fragment = start_html + len(prefix_bytes)
    end_fragment = start_fragment + len(html_bytes)
    end_html = end_fragment + len(suffix_bytes)
    
    header_str = header.format(start_html, end_html, start_fragment, end_fragment)
    clipboard_data = header_str.encode('utf-8') + prefix_bytes + html_bytes + suffix_bytes
    
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.SetClipboardData(cf_html, clipboard_data)
        print("Clipboard set!")
    finally:
        win32clipboard.CloseClipboard()

copy_html_to_clipboard("<h1>Hello World</h1><table border='1'><tr><td>Cell 1</td></tr></table>")
