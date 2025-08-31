# PowerShell script to search RSS feeds for keywords and generate a responsive HTML report

# Define keywords
$keywords = @(
    "Microsoft 365",
    "Defender",
    "Azure",
    "AzureAD",
    "Azure Active Directory",
    "EntraID",
    "Entra ID",
    "PIM",
    "SharePoint Online",
    "ExchangeOnline",
    "Intune",
    "Copilot"
)

# Define RSS feed URLs
$rssFeeds = @(
    'https://blogs.microsoft.com/feed/',
    'https://www.microsoft.com/en-us/education/blog/feed/',
    'https://news.microsoft.com/feed/',
    'http://feeds.windowscentral.com/wmexperts',
    'https://blogs.windows.com/feed/',
    'https://mspoweruser.com/feed/',
    'https://devblogs.microsoft.com/commandline/feed/',
    'https://msftnewsnow.com/feed/',
    'https://feeds.feedburner.com/perficient/microsoft',
    'https://www.thurrott.com/windows/feed',
    'https://www.thetraininglady.com/feed/',
    'https://www.metaoption.com/blog/feed/',
    'https://dellenny.com/feed/',
    'https://www.techbubbles.com/feed/',
    'https://hypervlab.co.uk/?feed=rss2',
    'https://abouconde.com/feed/',
    'https://www.zdnet.com/topic/microsoft/rss.xml',
    'https://techcommunity.microsoft.com/t5/s/gxcuf89792/rss/Community',
    'https://winbuzzer.com/feed/',
    'https://jussiroine.com/feed/',
    'https://www.windowslatest.com/feed/'
)

# Prepare results array
$results = @()

foreach ($feed in $rssFeeds) {
    try {
        $rss = [xml](Invoke-WebRequest -Uri $feed -UseBasicParsing -ErrorAction Stop).Content
        $items = $rss.rss.channel.item
        if (-not $items) { $items = $rss.feed.entry } # Atom support
        foreach ($item in $items) {
            $title = $item.title -as [string]
            $desc = $item.description -as [string]
            if (-not $desc) { $desc = $item.summary -as [string] }
            $link = $item.link
            if ($link -is [System.Xml.XmlElement]) { $link = $link.href }
                # Handle title and description fields that may be XML elements
                if ($item.title -is [System.Xml.XmlElement]) {
                    $title = $item.title.InnerText
                } else {
                    $title = $item.title -as [string]
                }
                if ($item.description -is [System.Xml.XmlElement]) {
                    $desc = $item.description.InnerText
                } elseif ($item.summary -is [System.Xml.XmlElement]) {
                    $desc = $item.summary.InnerText
                } else {
                    $desc = $item.description -as [string]
                    if (-not $desc) { $desc = $item.summary -as [string] }
                }

            # Get publish date (RSS: pubDate, Atom: published/updated)
            $pubDate = $null
            if ($item.pubDate) { $pubDate = Get-Date $item.pubDate -ErrorAction SilentlyContinue }
            elseif ($item.published) { $pubDate = Get-Date $item.published -ErrorAction SilentlyContinue }
            elseif ($item.updated) { $pubDate = Get-Date $item.updated -ErrorAction SilentlyContinue }

            # Only include articles from the last 7 days
            if ($pubDate -and $pubDate -lt (Get-Date).AddDays(-7)) { continue }

            $match = $false
            foreach ($kw in $keywords) {
                if ($title -match [regex]::Escape($kw) -or $desc -match [regex]::Escape($kw)) {
                    $match = $true
                    break
                }
            }
            if ($match) {
                $results += [PSCustomObject]@{
                    Title = $title
                    Link = $link
                    Description = $desc
                    Feed = $feed
                    Published = $pubDate
                }
            }
        }
    } catch {
    Write-Warning ("Failed to process {0}: {1}" -f $feed, $_)
    }
}

# Generate responsive HTML report

# Calculate date range for the report
$minDate = ($results | Where-Object { $_.Published } | Sort-Object Published | Select-Object -First 1).Published
$maxDate = ($results | Where-Object { $_.Published } | Sort-Object Published -Descending | Select-Object -First 1).Published
$dateRange = if ($minDate -and $maxDate) {
    "Articles from $($minDate.ToString('yyyy-MM-dd')) to $($maxDate.ToString('yyyy-MM-dd'))"
} else {
    "No articles found in the last 7 days."
}

# Count the number of articles
$articleCount = $results.Count

$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Topic Aggregator Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #9caf88; color: #111; }
        h1, h2, h3, h4, h5, h6, p, small { color: #111; }
    .card { background: #111; color: #fff; margin: 1em 0; padding: 1em; border-radius: 8px; box-shadow: 0 2px 8px #0001; }
    .card a { color: #8ecae6; text-decoration: none; }
    .card a:hover { text-decoration: underline; }
    .card h2, .card p, .card small { color: #fff; }
        @media (max-width: 600px) {
            body { margin: 5px; }
            .card { padding: 0.5em; }
        }
    </style>
</head>
<body>
    <h1>Topic Aggregator Report</h1>
    <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
    <p><strong>$dateRange</strong></p>
    <p><em>Number of articles: $articleCount</em></p>
    <div>
"@

foreach ($result in $results) {
    $pub = if ($result.Published) { $result.Published.ToString('yyyy-MM-dd') } else { '' }
    $html += @"
        <div class='card'>
            <h2>$(($result.Title -replace "<.*?>", ""))</h2>
            <p>$(($result.Description -replace "<.*?>", "") -replace "[\r\n]+", " ")</p>
            <a href='$(($result.Link -replace "'", "&apos;"))' target='_blank' style='display:inline-block;margin:0.5em 0;padding:0.5em 1em;background:#0078d4;color:#fff;border:none;border-radius:4px;text-decoration:none;font-weight:bold;'>Read Article</a><br/>
            <small>Feed: $($result.Feed) | Published: $pub</small>
        </div>
"@
}

$html += @"
    </div>
</body>
</html>
"@

# Save to HTML file
$reportPath = Join-Path $PSScriptRoot 'DailyTopicReport.html'
$html | Set-Content -Path $reportPath -Encoding UTF8
Write-Host "Report saved to $reportPath"
