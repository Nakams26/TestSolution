/**
 * Claude PR Reviewer
 * Reviews pull requests using Claude and auto-merges if approved.
 */

const Anthropic = require('@anthropic-ai/sdk');
const fetch = require('node-fetch');

const GITHUB_TOKEN  = process.env.GITHUB_TOKEN;
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY;
const REPO          = process.env.GITHUB_REPOSITORY;       // "owner/repo"
const PR_NUMBER     = process.env.PR_NUMBER;

const github = (path, options = {}) =>
  fetch(`https://api.github.com/repos/${REPO}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      ...options.headers,
    },
  });

async function getPRDiff() {
  const res = await fetch(`https://api.github.com/repos/${REPO}/pulls/${PR_NUMBER}`, {
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github.v3.diff',
      'X-GitHub-Api-Version': '2022-11-28',
    },
  });
  return res.text();
}

async function getPRInfo() {
  const res = await github(`/pulls/${PR_NUMBER}`);
  return res.json();
}

async function postReview(verdict, body) {
  const event = verdict === 'APPROVE' ? 'APPROVE' : 'REQUEST_CHANGES';
  const res = await github(`/pulls/${PR_NUMBER}/reviews`, {
    method: 'POST',
    body: JSON.stringify({ body, event }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Failed to post review: ${JSON.stringify(data)}`);
  console.log(`Review posted: ${event}`);
  return data;
}

async function mergePR(title) {
  const res = await github(`/pulls/${PR_NUMBER}/merge`, {
    method: 'PUT',
    body: JSON.stringify({
      commit_title: title,
      merge_method: 'squash',
    }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Failed to merge PR: ${JSON.stringify(data)}`);
  console.log('PR merged successfully');
}

async function main() {
  const client = new Anthropic({ apiKey: ANTHROPIC_KEY });

  const [diff, prInfo] = await Promise.all([getPRDiff(), getPRInfo()]);

  console.log(`Reviewing PR #${PR_NUMBER}: ${prInfo.title}`);

  const response = await client.messages.create({
    model: 'claude-opus-4-6',
    max_tokens: 4096,
    system: `You are an expert reviewer for a Power Platform / Dataverse solution managed as code.
The solution consists of:
- XML files in exported_unpacked/ (entity definitions, forms, views, relationships)
- Node.js scripts for deployment (create_env_vars.js, import_solution.js)
- GitHub Actions workflows in .github/workflows/

Your job is to review pull requests for correctness and safety. Focus on:
1. Valid XML structure for Dataverse entities (correct types, required fields)
2. No breaking changes (e.g., removing existing columns that may have data)
3. Correct column types (money, decimal, datetime, string, picklist, etc.)
4. No security issues in scripts

Respond in this exact format:
VERDICT: APPROVE or REQUEST_CHANGES
SUMMARY: (1-2 sentence summary of what the PR does)
DETAILS: (bullet points of findings — issues if REQUEST_CHANGES, confirmations if APPROVE)`,
    messages: [
      {
        role: 'user',
        content: `PR Title: ${prInfo.title}\nPR Description: ${prInfo.body || '(none)'}\n\nDiff:\n\`\`\`diff\n${diff}\n\`\`\``,
      },
    ],
  });

  const review = response.content[0].text;
  console.log('\nClaude review:\n', review);

  const verdictMatch = review.match(/VERDICT:\s*(APPROVE|REQUEST_CHANGES)/);
  if (!verdictMatch) throw new Error('Could not parse verdict from Claude response');

  const verdict = verdictMatch[1];
  await postReview(verdict, review);

  if (verdict === 'APPROVE') {
    await mergePR(prInfo.title);
  } else {
    console.log('PR has issues — not merging.');
    process.exit(1);
  }
}

main().catch(e => { console.error(e.message); process.exit(1); });
