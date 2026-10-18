import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import { copyFile } from "wpsjs/vite_plugins"

// https://vitejs.dev/config/
export default defineConfig({
    base: './',
    plugins: [
        copyFile({ src: 'manifest.xml', dest: 'manifest.xml', }),
        copyFile({ src: 'main.js', dest: 'main.js', }),
        copyFile({ src: 'ribbon.xml', dest: 'ribbon.xml', }),
        copyFile({ src: 'taskpane.html', dest: 'taskpane.html', }),
        copyFile({ src: 'taskpane.css', dest: 'taskpane.css', }),
        copyFile({ src: 'taskpane.js', dest: 'taskpane.js', }),
        copyFile({ src: 'app.js', dest: 'app.js', }),
        copyFile({ src: 'modules', dest: 'modules', }),
        copyFile({ src: 'lib', dest: 'lib', }),
    ],
    resolve: {
        alias: {
            '@': fileURLToPath(new URL('./src', import.meta.url))
        }
    },
    server: {
        host: '0.0.0.0'
    }
})
