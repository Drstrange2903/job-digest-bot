# AI/ML Job Digest Bot

This repo runs a scheduled job-search digest for Vraj Patel.

## What It Does

- Searches public job sources and useful search URLs for AI/ML, GenAI, RAG, LLM, Data Scientist, and Data Engineer roles.
- Filters for preferred locations: Bangalore/Bengaluru, Pune, Hyderabad, Mumbai, or Remote/Hybrid.
- Excludes Ahmedabad, Vadodara, Gurgaon/Gurugram, and Noida unless the role is fully remote.
- Builds an Excel tracker.
- Emails the digest and tracker attachment.

## GitHub Actions Schedule

Current test schedule: every 10 minutes.

After confirming email delivery, change `.github/workflows/job-digest.yml` from:

```yaml
- cron: "*/10 * * * *"
```

to every 2 hours:

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
