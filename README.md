# Company Research Agent

An N8N workflow that orchestrates multi-source web research and AI synthesis to produce structured intelligence packets ready for import into Google NotebookLM.

## What it does

Given a company name, a job posting URL, and a seniority level (PM / Senior PM / Staff PM / Lead PM), the workflow:
1. Fetches data from multiple web sources in parallel
2. Synthesizes findings into a structured research packet
3. Formats the output for direct import into NotebookLM

## Stack

- **N8N** — workflow orchestration and scheduling
- **Google NotebookLM** — AI-powered synthesis layer
- **Web APIs** — multi-source data ingestion

## Usage

1. Import the workflow JSON from `Company-Research-Agent/latest version/` into your N8N instance
2. Fill in the form: company name, role URL, and seniority level
3. Import the output packet into NotebookLM

## Version history

| File | Description |
|---|---|
| `latest version/PM Interview Company Research Agent-latest.json` | Current active version |
| `versions/PM Interview Company Research Agent-v4.json` | Latest numbered release |
| `versions/PM Interview Company Research Agent-v3.json` | Previous version |
| `versions/Company Research - Source Aggregator-v2.json` | Companion source aggregator workflow |
