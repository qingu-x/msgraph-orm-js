import { defineConfig } from 'vite'
import { resolve } from 'path'
import dts from 'vite-plugin-dts'
import { builtinModules } from 'module'
import { writeFileSync, readFileSync } from 'fs'

export default defineConfig({
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'MSGRAPH_ORM_JS',
      formats: ['es', 'cjs', 'umd'],
      fileName: (format) => {
        if (format === 'es') return 'bundle.esm.js'
        if (format === 'cjs') return 'bundle.cjs.js'
        if (format === 'umd') return 'bundle.browser.js'
        return `bundle.${format}.js`
      }
    },
    outDir: 'lib',
    rollupOptions: {
      external: (id) => {
        // NPM 包始终为外部依赖
        const npmExternals = [
          '@azure/identity',
          '@microsoft/microsoft-graph-client',
          '@microsoft/microsoft-graph-types',
          'https-proxy-agent',  // Node.js 专用代理包
          'agent-base'          // https-proxy-agent 的依赖
        ]
        
        if (npmExternals.some(pkg => id === pkg || id.startsWith(pkg + '/'))) {
          return true
        }
        
        // Node.js 内置模块为外部依赖
        const nodeBuiltins = builtinModules
        if (nodeBuiltins.includes(id) || nodeBuiltins.includes(id.replace(/^node:/, ''))) {
          return true
        }
        
        return false
      },
      output: {
        globals: {
          '@azure/identity': 'AzureIdentity',
          '@microsoft/microsoft-graph-client': 'MicrosoftGraph',
          '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js': 'MicrosoftGraphAuthProviders',
          '@microsoft/microsoft-graph-types': 'MicrosoftGraphTypes',
          'https-proxy-agent': 'HttpsProxyAgent'
        }
      }
    },
    sourcemap: true,
    minify: 'terser',
    terserOptions: {
      compress: {
        drop_console: false,
      }
    }
  },
  plugins: [
    dts({
      outDir: 'types',
      insertTypesEntry: true,
      copyDtsFiles: true,
      include: ['src/**/*'],
      exclude: ['test/**/*', 'node_modules/**/*']
    }),
    {
      name: 'fix-browser-bundle',
      closeBundle() {
        console.log('Running fix-browser-bundle plugin...');
        // 修复浏览器打包文件，添加正确的 IIFE 结构
        const browserBundlePath = resolve(__dirname, 'lib/bundle.browser.js');
        try {
          const content = readFileSync(browserBundlePath, 'utf-8');
          console.log('Read browser bundle content, length:', content.length);
          
          // 强制执行修复
          console.log('Adding IIFE structure to browser bundle...');
          // 移除可能存在的部分IIFE结构和source mapping URL
          let cleanContent = content.trim();
          if (cleanContent.startsWith('!function(')) {
            // 如果已经有部分IIFE结构，尝试移除它
            const iifeStart = cleanContent.indexOf('{"use strict";');
            if (iifeStart > 0) {
              cleanContent = cleanContent.substring(iifeStart + 13); // 跳过{"use strict";
            }
          }
          // 移除source mapping URL
          cleanContent = cleanContent.replace(/;\s*\/\/#\s*sourceMappingURL=bundle\.browser\.js\.map\s*$/, '');
          
          // 添加完整的 IIFE 包装
          const fixedContent = `!function(e,t){"object"==typeof exports&&"undefined"!=typeof module?t(exports,require("@azure/identity"),require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js"),require("https-proxy-agent"),require("@microsoft/microsoft-graph-client")):"function"==typeof define&&define.amd?define(["exports","@azure/identity","@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js","https-proxy-agent","@microsoft/microsoft-graph-client"],t):t((e="undefined"!=typeof globalThis?globalThis:e||self).MSGRAPH_ORM_JS={},e.AzureIdentity,e.MicrosoftGraphAuthProviders,e.HttpsProxyAgent,e.MicrosoftGraph)}(this,function(e,t,s,r,n){"use strict";` + 
            cleanContent + 
            `;return e;});\n//# sourceMappingURL=bundle.browser.js.map`;
          writeFileSync(browserBundlePath, fixedContent);
          console.log('Fixed browser bundle with IIFE structure');
        } catch (error) {
          console.error('Failed to fix browser bundle:', error);
        }
      }
    }
  ],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src')
    }
  }
})
