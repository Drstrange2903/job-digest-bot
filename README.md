# AI/ML Job Digest Bot

This repo runs a scheduled job-search digest for Vraj Patel.

## What It Does

- Searches public job sources, browser-rendered Naukri search result pages, and useful search URLs for AI/ML, GenAI, RAG, LLM, Data Scientist, and Data Engineer roles.
- Filters for preferred locations: Bangalore/Bengaluru, Pune, Hyderabad, Mumbai, or Remote/Hybrid.
- Excludes Ahmedabad, Vadodara, Gurgaon/Gurugram, and Noida unless the role is fully remote.
- Builds an Excel tracker.
- Emails the digest and tracker attachment.
- Remembers already-sent job links between scheduled runs, so it does not keep emailing the same jobs.

## Naukri / LinkedIn Accounts

LinkedIn and Naukri alerts are configured in the browser account directly. The bot does not store portal passwords or session cookies in GitHub.

The GitHub bot uses public Naukri search pages based on the same preferences:

- AI ML Engineer Python LLM
- Generative AI Engineer RAG Python
- Machine Learning Engineer Python
- Data Scientist Python Machine Learning
- Data Engineer Python SQL AI

## GitHub Actions Schedule

Current production schedule:

```yaml
- cron: "0 */2 * * *"
```

## Required GitHub Secrets

Set these in GitHub repo settings:

- `SMTP_PASSWORD`: Gmail app password for `patelvraj09876@gmail.com`, not the normal Gmail password

The workflow is currently hardcoded to send from and to:

```text
patelvraj09876@gmail.com
```

## Manual Local Run

```powershell
python -m pip install -r requirements.txt
python job_digest_bot.py
```

To send email locally, set the SMTP environment variables first.
