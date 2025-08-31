
<#
.SYNOPSIS
    Aggregates articles from multiple RSS feeds, filters by Microsoft-related keywords, and generates a responsive HTML report.
    Optionally, summarizes each article using Azure OpenAI GPT-4.

.DESCRIPTION
    This script fetches articles from a list of RSS feeds, filters them by a set of Microsoft/cloud/security keywords, and generates a modern, accessible HTML report. If Azure OpenAI parameters are provided and summarization is enabled, the script submits the full content of each article to Azure OpenAI for a thoughtful, technical summary using GPT-4.

.PARAMETER AzureOpenAIEndpoint
    The endpoint URL for your Azure OpenAI resource (e.g., https://<resource-name>.openai.azure.com).

.PARAMETER AzureOpenAIApiKey
    The API key for your Azure OpenAI resource.

.PARAMETER AzureOpenAIDeployment
    The deployment name for your GPT-4 model in Azure OpenAI.

.PARAMETER SummarizeWithOpenAI
    Switch to enable summarization of articles using Azure OpenAI.

.NOTES
    Author: wirelesshobo
    Created: 2024-08-31
    Last Updated: 2024-08-31
    GitHub: https://github.com/wirelesshobo/Group-TopicAggregator

.EXAMPLE
    .\Group-TopicAggregator.ps1
    Runs the script and generates a report with summaries/links for recent Microsoft-related articles.

.EXAMPLE
    .\Group-TopicAggregator.ps1 -AzureOpenAIEndpoint "https://myopenai.openai.azure.com" -AzureOpenAIApiKey "<key>" -AzureOpenAIDeployment "gpt-4" -SummarizeWithOpenAI
    Runs the script and generates a report with AI-generated summaries for each article.
#>

param(
    [string]$AzureOpenAIEndpoint = '',
    [string]$AzureOpenAIApiKey = '',
    [string]$AzureOpenAIDeployment = '',
    [switch]$SummarizeWithOpenAI
)


<#
.SYNOPSIS
    Summarizes text using Azure OpenAI GPT-4 via REST API.
.DESCRIPTION
    Sends the provided content to the specified Azure OpenAI deployment and returns the summary.
.PARAMETER Content
    The article content to summarize.
.PARAMETER Endpoint
    The Azure OpenAI endpoint URL.
.PARAMETER ApiKey
    The Azure OpenAI API key.
.PARAMETER Deployment
    The Azure OpenAI deployment name.
.OUTPUTS
    [string] The summary or error message.
#>
function Get-OpenAISummary {
    param(
        [string]$Content,
        [string]$Endpoint,
        [string]$ApiKey,
        [string]$Deployment
    )
    $uri = "$Endpoint/openai/deployments/$Deployment/chat/completions?api-version=2024-02-15-preview"
    $headers = @{ 'api-key' = $ApiKey; 'Content-Type' = 'application/json' }
    $body = @{ 
        messages = @(
            @{ role = 'system'; content = 'You are an expert technical writer. Summarize the following article in a thoughtful, meaningful, and concise way for a technical audience.' },
            @{ role = 'user'; content = $Content }
        )
        max_tokens = 512
        temperature = 0.3
        top_p = 1
        frequency_penalty = 0
        presence_penalty = 0
    } | ConvertTo-Json -Depth 5
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $body
        return $response.choices[0].message.content
    } catch {
        $errMsg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $errMsg += " | Details: $($_.ErrorDetails.Message)"
        }
        return "[Summary unavailable: $errMsg]"
    }
}


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

$rssFeeds = @(
    'https://blogs.microsoft.com/feed/',
    'https://www.microsoft.com/en-us/education/blog/feed/',
    'https://dellenny.com/feed/'
)

# Prepare results array
$results = @()

foreach ($feed in $rssFeeds) {
    try {
        # Download and parse RSS feed
        $rss = [xml](Invoke-WebRequest -Uri $feed -UseBasicParsing -ErrorAction Stop).Content
        $items = $rss.rss.channel.item
        if (-not $items) { $items = $rss.feed.entry } # Atom support
        foreach ($item in $items) {
            # Extract title and description, handling XML elements
            $title = $item.title -as [string]
            $desc = $item.description -as [string]
            if (-not $desc) { $desc = $item.summary -as [string] }
            $link = $item.link
            if ($link -is [System.Xml.XmlElement]) { $link = $link.href }
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

            # Check for keyword match
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
                $summary = ''
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
                        if (-not $fullContent) { $fullContent = '[No main content found]' }
                        # Summarize with Azure OpenAI if enabled and parameters provided
                        if ($SummarizeWithOpenAI -and $AzureOpenAIEndpoint -and $AzureOpenAIApiKey -and $AzureOpenAIDeployment) {
                            $summary = Get-OpenAISummary -Content $fullContent -Endpoint $AzureOpenAIEndpoint -ApiKey $AzureOpenAIApiKey -Deployment $AzureOpenAIDeployment
                        }
                    }
                } catch {
                    $fullContent = '[Could not retrieve full content]'
                    $summary = ''
                }
                $results += [PSCustomObject]@{
                    Title = $title
                    Link = $link
                    Description = $desc
                    Feed = $feed
                    Published = $pubDate
                    FullContent = $fullContent
                    Summary = $summary
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
    $summaryHtml = ''
    if ($result.Summary) {
        $summaryHtml = "<div style='margin:0.5em 0; padding:0.5em; background:#e6f2ff; color:#111; border-left:4px solid #005A9E;'><strong>AI Summary:</strong> $($result.Summary)</div>"
    }
    $html += @"
        <div class='card'>
            <h2>$(($result.Title -replace "<.*?>", ""))</h2>
            <p>$(($result.Description -replace "<.*?>", "") -replace "[\r\n]+", " ")</p>
            $summaryHtml
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
