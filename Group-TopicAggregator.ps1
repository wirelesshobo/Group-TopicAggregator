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
                # Try to fetch the full article content (best effort)
                $fullContent = ''
                try {
                    if ($link -and $link -is [string] -and $link.StartsWith('http')) {
                        $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36' }
                        $articleResponse = Invoke-WebRequest -Uri $link -UseBasicParsing -TimeoutSec 20 -Headers $headers -ErrorAction Stop
                        $htmlContent = $articleResponse.Content
                        # Try to extract <article>, <main>, or largest <div> using regex
                        $fullContent = ''
                        if ($htmlContent -match '<article[\s\S]*?</article>') {
                            $fullContent = $Matches[0]
                        } elseif ($htmlContent -match '<main[\s\S]*?</main>') {
                            $fullContent = $Matches[0]
                        } else {
                            # Find all divs and pick the largest by text length
                            $divMatches = [regex]::Matches($htmlContent, '<div[\s\S]*?</div>')
                            $maxLen = 0
                            foreach ($divMatch in $divMatches) {
                                $divText = $divMatch.Value -replace '<.*?>', ''
                                if ($divText.Length -gt $maxLen) {
                                    $maxLen = $divText.Length
                                    $fullContent = $divMatch.Value
                                }
                            }
                        }
                        # Clean up HTML tags
                        $fullContent = $fullContent -replace '<.*?>', ''
                        # Do not truncate the full content
                        if (-not $fullContent) { $fullContent = '[No main content found]' }
                    }
                } catch {
                    $fullContent = '[Could not retrieve full content]'
                }
                $results += [PSCustomObject]@{
                    Title = $title
                    Link = $link
                    Description = $desc
                    Feed = $feed
                    Published = $pubDate
                    FullContent = $fullContent
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
        body { font-family: Arial, sans-serif; margin: 20px; background: #fff; color: #111; }
        h1, h2, h3, h4, h5, h6 { color: #005A9E; }
        p, small { color: #333; }
        .card {
            background: #222;
            color: #fff;
            margin: 1em 0;
            padding: 1em;
            border-radius: 8px;
            box-shadow: 0 2px 8px #0001;
            border: 1px solid #ccc;
        }
        .card a {
            color: #0066CC;
            text-decoration: none;
            background: #0078D4;
            color: #fff;
            border: none;
            border-radius: 4px;
            padding: 0.5em 1em;
            font-weight: bold;
            display: inline-block;
            margin: 0.5em 0;
        }
        .card a:hover {
            background: #004A99;
            color: #fff;
            text-decoration: underline;
        }
        .card h2, .card p, .card small {
            color: #fff;
        }
        details summary {
            color: #005A9E;
            cursor: pointer;
            font-weight: bold;
        }
        details[open] summary {
            color: #004A99;
        }
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
    $fullContentHtml = ''
    if ($result.FullContent) {
        $cleanContent = ($result.FullContent -replace "<.*?>", "") -replace "[\r\n]+", " "
        $fullContentHtml = "<details><summary>Show Full Content</summary><div style='margin-top:0.5em;'>$cleanContent</div></details>"
    }
    $html += @"
        <div class='card'>
            <h2>$(($result.Title -replace "<.*?>", ""))</h2>
            <p>$(($result.Description -replace "<.*?>", "") -replace "[\r\n]+", " ")</p>
            <a href='$(($result.Link -replace "'", "&apos;"))' target='_blank' style='display:inline-block;margin:0.5em 0;padding:0.5em 1em;background:#0078d4;color:#fff;border:none;border-radius:4px;text-decoration:none;font-weight:bold;'>Read Article</a><br/>
            $fullContentHtml
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
