// Konfigurasi Nuxt 3 untuk blog AdSense static generation
export default defineNuxtConfig({
  ssr: true,
  modules: [
    '@nuxt/content',
    '@nuxtjs/tailwindcss'
  ],
  app: {
    head: {
      htmlAttrs: {
        lang: 'id'
      },
      title: 'KerjaKantor - Panduan Excel & Google Sheets Terlengkap',
      meta: [
        { charset: 'utf-8' },
        { name: 'viewport', content: 'width=device-width, initial-scale=1' },
        { name: 'description', content: 'Panduan praktis rumus Excel, Google Sheets, otomasi spreadsheet, dan template analisis data untuk produktivitas kerja.' },
        // Meta tag AdSense account
        { name: 'google-adsense-account', content: 'ca-pub-8340335305046609' }
      ],
      link: [
        { rel: 'icon', type: 'image/x-icon', href: '/favicon.ico' }
      ],
      script: [
        // AdSense script loader
        {
          src: 'https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-8340335305046609',
          async: true,
          crossorigin: 'anonymous'
        }
      ]
    }
  },
  content: {
    highlight: {
      theme: 'github-dark'
    }
  },
  nitro: {
    preset: 'cloudflare_module',
    prerender: {
      routes: ['/sitemap.xml', '/robots.txt']
    }
  },
  vite: {
    server: {
      allowedHosts: true
    }
  }
})
