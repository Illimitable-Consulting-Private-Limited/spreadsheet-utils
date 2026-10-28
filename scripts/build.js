import esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

const distDir = path.resolve('dist');

async function clean() {
    if (fs.existsSync(distDir)) {
        fs.rmSync(distDir, { recursive: true, force: true });
    }
    fs.mkdirSync(path.join(distDir, 'esm'), { recursive: true });
    fs.mkdirSync(path.join(distDir, 'cjs'), { recursive: true });
    fs.mkdirSync(path.join(distDir, 'browser'), { recursive: true });
}

async function build() {
    console.log('Building spreadsheet-utils multi-target distribution...\n');
    await clean();

    const emptyShimPath = path.resolve('src/utils/empty-shim.js');
    const universalUnpackPath = path.resolve('src/export/unpackXlsxFileUniversal.js');

    const browserShimPlugin = {
        name: 'browser-shim',
        setup(build) {
            build.onResolve({ filter: /unpackXlsxFileNode\.js$/ }, () => {
                return { path: universalUnpackPath };
            });
        }
    };

    // 1. ESM Builds (Node / Modern Bundlers)
    console.log('1. Generating ESM builds (dist/esm/)...');
    await esbuild.build({
        entryPoints: [
            { in: 'src/index.js', out: 'index' },
            { in: 'src/plugins/index.js', out: 'plugins' },
            { in: 'src/write/index.js', out: 'write' },
            { in: 'src/full.js', out: 'full' },
            { in: 'src/worker/index.js', out: 'worker' }
        ],
        outdir: 'dist/esm',
        outExtension: { '.js': '.mjs' },
        format: 'esm',
        platform: 'node',
        target: 'node18',
        bundle: true,
        packages: 'external',
        sourcemap: true
    });

    // 2. CommonJS Builds (Node require)
    console.log('2. Generating CommonJS builds (dist/cjs/)...');
    await esbuild.build({
        entryPoints: [
            { in: 'src/index.js', out: 'index' },
            { in: 'src/plugins/index.js', out: 'plugins' },
            { in: 'src/write/index.js', out: 'write' },
            { in: 'src/full.js', out: 'full' },
            { in: 'src/worker/index.js', out: 'worker' }
        ],
        outdir: 'dist/cjs',
        outExtension: { '.js': '.cjs' },
        format: 'cjs',
        platform: 'node',
        target: 'node18',
        bundle: true,
        packages: 'external',
        sourcemap: true
    });

    // 3. Browser Standalone IIFE Bundles (for direct <script> tags)
    console.log('3. Generating Browser standalone bundles (dist/browser/)...');
    const browserConfigs = [
        {
            in: 'src/index.js',
            out: 'dist/browser/spreadsheet-utils.min.js',
            globalName: 'spreadsheetUtils'
        },
        {
            in: 'src/plugins/index.js',
            out: 'dist/browser/spreadsheet-utils-plugins.min.js',
            globalName: 'spreadsheetUtilsPlugins'
        },
        {
            in: 'src/write/browser.js',
            out: 'dist/browser/spreadsheet-utils-write.min.js',
            globalName: 'spreadsheetUtilsWrite'
        },
        {
            in: 'src/full-browser.js',
            out: 'dist/browser/spreadsheet-utils-full.min.js',
            globalName: 'spreadsheetUtilsFull'
        }
    ];

    for (const cfg of browserConfigs) {
        await esbuild.build({
            entryPoints: [cfg.in],
            outfile: cfg.out,
            format: 'iife',
            globalName: cfg.globalName,
            platform: 'browser',
            target: 'es2020',
            bundle: true,
            minify: true,
            sourcemap: true,
            plugins: [browserShimPlugin],
            define: {
                'process.env.NODE_ENV': '"production"'
            },
            alias: {
                'fs': emptyShimPath,
                'node:fs': emptyShimPath,
                'stream': emptyShimPath,
                'node:stream': emptyShimPath,
                'buffer': emptyShimPath,
                'node:buffer': emptyShimPath,
                'path': emptyShimPath,
                'node:path': emptyShimPath,
                'zlib': emptyShimPath,
                'node:zlib': emptyShimPath,
                'unzipper-esm': emptyShimPath
            }
        });
    }

    console.log('\nBuild completed successfully! Artifacts generated in dist/:');
    console.log('  dist/esm/:     index.mjs, plugins.mjs, write.mjs, full.mjs');
    console.log('  dist/cjs/:     index.cjs, plugins.cjs, write.cjs, full.cjs');
    console.log('  dist/browser/: spreadsheet-utils.min.js, spreadsheet-utils-plugins.min.js, spreadsheet-utils-write.min.js, spreadsheet-utils-full.min.js');
}

build().catch((err) => {
    console.error('Build failed:', err);
    process.exit(1);
});
