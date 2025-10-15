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
        },
        // 明确指定UMD格式的配置
        format: 'umd'
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
          
          // 始终执行修复以确保正确的结构
          console.log('Fixing browser bundle structure...');
          // 清理内容，移除可能的错误结构
          let cleanContent = content.trim();
          
          // 如果内容以IIFE开头，需要提取内部内容
          if (cleanContent.startsWith('!function(')) {
            // 查找"use strict"的位置
            const useStrictIndex = cleanContent.indexOf('"use strict";');
            if (useStrictIndex > 0) {
              // 提取"use strict";之后的内容
              cleanContent = cleanContent.substring(useStrictIndex + 13);
            } else {
              // 查找模块主体内容的开始位置
              const moduleBodyStart = cleanContent.indexOf('{');
              if (moduleBodyStart > 0) {
                // 找到第一个{之后的内容
                const actualContentStart = cleanContent.indexOf('"use strict";', moduleBodyStart);
                if (actualContentStart > 0) {
                  cleanContent = cleanContent.substring(actualContentStart + 13);
                } else {
                  // 尝试查找var关键字作为内容开始
                  const varStart = cleanContent.indexOf('var ', moduleBodyStart);
                  if (varStart > 0) {
                    cleanContent = cleanContent.substring(varStart);
                  }
                }
              }
            }
          }
          
          // 确保清理掉末尾的source mapping URL和任何多余的括号
          cleanContent = cleanContent.replace(/;\s*\/\/#\s*sourceMappingURL=bundle\.browser\.js\.map\s*$/, '');
          cleanContent = cleanContent.replace(/\s*}\);\s*$/, ''); // 移除可能的结尾括号
          cleanContent = cleanContent.replace(/\n*$/, ''); // 移除末尾的换行符
          
          // 构建正确的IIFE结构，包含source mapping URL
          const fixedContent = `!function(e,t){"object"==typeof exports&&"undefined"!=typeof module?t(exports,require("@azure/identity"),require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js"),require("https-proxy-agent"),require("@microsoft/microsoft-graph-client")):"function"==typeof define&&define.amd?define(["exports","@azure/identity","@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js","https-proxy-agent","@microsoft/microsoft-graph-client"],t):t((e="undefined"!=typeof globalThis?globalThis:e||self).MSGRAPH_ORM_JS={},e.AzureIdentity,e.MicrosoftGraphAuthProviders,e.HttpsProxyAgent,e.MicrosoftGraph)}(this,function(e,t,s,r,n){"use strict";` + 
            cleanContent + 
            `;return e;});\n//# sourceMappingURL=bundle.browser.js.map`;
            
          writeFileSync(browserBundlePath, fixedContent);
          console.log('Fixed browser bundle with correct IIFE structure');
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
