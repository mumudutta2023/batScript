$folder = [uri]'C:\PracPull'
foreach ($w in (New-Object -ComObject Shell.Application).Windows()) {
    if ($w.LocationURL -ieq $folder.AbsoluteUri) { $w.Quit(); break }
}