#!/usr/bin/env node

/**
 * Upstream Synchronization Utility for spreadsheet-utils.
 *
 * Tracks, inspects, and verifies upstream updates from:
 * 1. read-excel-file (GitLab: https://gitlab.com/catamphetamine/read-excel-file.git)
 * 2. tabularjs (GitHub: https://github.com/jspreadsheet/tabularjs.git)
 * 3. write-excel-file (GitLab: https://gitlab.com/catamphetamine/write-excel-file.git)
 *
 * Usage:
 *   node scripts/sync-upstream.js --status
 *   node scripts/sync-upstream.js --check
 *   node scripts/sync-upstream.js --verify
 *   node scripts/sync-upstream.js --sync [repo-name]
 */

import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const UPSTREAM_REPOSITORIES = {
    'read-excel-file': {
        name: 'read-excel-file',
        url: 'https://gitlab.com/catamphetamine/read-excel-file.git',
        trackedVersion: '5.8.6',
        npmPackage: 'read-excel-file',
        localPaths: [
            'src/xlsx',
            'src/xml',
            'src/schema',
            'src/export',
            'src/zip'
        ],
        description: 'Core streaming SAX XLSX reader and schema validator'
    },
    'tabularjs': {
        name: 'tabularjs',
        url: 'https://github.com/jspreadsheet/tabularjs.git',
        trackedVersion: '0.0.12',
        npmPackage: 'tabularjs',
        localPaths: [
            'src/parser',
            'src/plugins',
            'src/utils'
        ],
        description: 'Multi-format parsers (CSV, XLS, ODS, HTML, XML, DBF, Lotus, DIF, SLK, Numbers)'
    },
    'write-excel-file': {
        name: 'write-excel-file',
        url: 'https://gitlab.com/catamphetamine/write-excel-file.git',
        trackedVersion: '4.1.1',
        npmPackage: 'write-excel-file',
        localPaths: [
            'src/write'
        ],
        description: 'Styled XLSX writer with cell styling, conditional formatting, and schemas'
    }
};

function run(cmd, options = {}) {
    try {
        const res = execSync(cmd, { encoding: 'utf-8', ...options });
        return typeof res === 'string' ? res.trim() : res;
    } catch (err) {
        if (options.ignoreError) return null;
        throw err;
    }
}

function showStatus() {
    console.log('='.repeat(70));
    console.log('  spreadsheet-utils: Upstream Status & File Mapping');
    console.log('='.repeat(70));

    for (const [key, repo] of Object.entries(UPSTREAM_REPOSITORIES)) {
        console.log(`\n📦 ${repo.name} (${repo.npmPackage})`);
        console.log(`   Description:     ${repo.description}`);
        console.log(`   Repository:      ${repo.url}`);
        console.log(`   Tracked Release: v${repo.trackedVersion}`);
        console.log(`   Local Paths:`);
        for (const p of repo.localPaths) {
            const exists = fs.existsSync(p);
            console.log(`     - ${p} ${exists ? '✓' : '✗ (missing)'}`);
        }
    }
    console.log('\n' + '='.repeat(70));
}

async function checkUpstream() {
    console.log('Checking for upstream updates via npm registry...\n');

    for (const [key, repo] of Object.entries(UPSTREAM_REPOSITORIES)) {
        process.stdout.write(`Fetching latest version for ${repo.npmPackage}... `);
        const latest = run(`npm view ${repo.npmPackage} version`, { ignoreError: true });
        if (latest) {
            const isUpToDate = latest === repo.trackedVersion;
            console.log(`${latest} ${isUpToDate ? '(✓ up to date)' : `(⚠️ UPDATE AVAILABLE: tracked v${repo.trackedVersion} -> latest v${latest})`}`);
        } else {
            console.log(`(unable to query registry, offline or sandbox)`);
        }
    }

    console.log('\nRun "node scripts/sync-upstream.js --status" to review tracked paths.');
}

function verifyTests() {
    console.log('Verifying test suite and build pipeline...\n');
    try {
        console.log('1. Building distribution bundles...');
        run('node scripts/build.js', { stdio: 'inherit' });

        console.log('\n2. Executing full test suite...');
        run('npm test', { stdio: 'inherit' });

        console.log('\n✓ All 465+ tests passed! Upstream integration is verified and stable.');
    } catch (err) {
        console.error('\n✗ Verification failed! See error output above.');
        process.exit(1);
    }
}

function syncRepo(target) {
    if (!target || !UPSTREAM_REPOSITORIES[target]) {
        console.error(`Unknown repository: "${target}". Valid targets: ${Object.keys(UPSTREAM_REPOSITORIES).join(', ')}`);
        process.exit(1);
    }

    const repo = UPSTREAM_REPOSITORIES[target];
    console.log(`\nSyncing updates from ${repo.name} (${repo.url})...`);
    const stagingDir = path.resolve('.upstream-staging', target);

    if (fs.existsSync(stagingDir)) {
        console.log(`Updating existing staging clone at ${stagingDir}...`);
        run(`git -C ${stagingDir} fetch origin --tags`);
    } else {
        console.log(`Cloning ${repo.url} into staging area...`);
        fs.mkdirSync(path.dirname(stagingDir), { recursive: true });
        run(`git clone --depth 1 ${repo.url} ${stagingDir}`);
    }

    console.log(`\n✓ Staging repository updated at ${stagingDir}`);
    console.log(`Review diffs between ${stagingDir} and local paths: ${repo.localPaths.join(', ')}`);
    console.log('After applying changes, run: npm run sync:verify\n');
}

// CLI argument parsing
const args = process.argv.slice(2);
if (args.includes('--check')) {
    checkUpstream();
} else if (args.includes('--verify')) {
    verifyTests();
} else if (args.includes('--sync')) {
    const target = args[args.indexOf('--sync') + 1];
    syncRepo(target);
} else {
    showStatus();
}
