import { defineConfig } from 'vite'
import { resolve } from 'path'
import dts from 'vite-plugin-dts'

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
      external: [
        '@azure/identity',
        '@microsoft/microsoft-graph-client',
        '@microsoft/microsoft-graph-types',
        'ky',
        'debug'
      ],
      output: {
        globals: {
          '@azure/identity': 'AzureIdentity',
          '@microsoft/microsoft-graph-client': 'MicrosoftGraph',
          '@microsoft/microsoft-graph-types': 'MicrosoftGraphTypes',
          'ky': 'ky',
          'debug': 'debug'
        },
        footer: (chunk) => {
          if (chunk.name === 'index') {
            return `
if(typeof window !== 'undefined') {
  window._msgraph_orm_js_version_ = '0.0.1'
}`
          }
          return ''
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
    })
  ],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src')
    }
  }
})
