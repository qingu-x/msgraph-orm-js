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
        // 确保可选链操作符被正确转换
        ecma: 2015
      },
      mangle: {
        // 保持类名和方法名不被混淆，确保兼容性
        keep_classnames: true,
        keep_fnames: true
      },
      format: {
        // 确保输出兼容性
        ecma: 2015
      }
    },
    target: 'es2015'
  },
  plugins: [
    dts({
      outDir: 'types',
      insertTypesEntry: true,
      copyDtsFiles: true,
      include: ['src/**/*'],
      exclude: ['test/**/*', 'node_modules/**/*']
    })
  ],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src')
    }
  }
})
