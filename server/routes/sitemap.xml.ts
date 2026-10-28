import { serverQueryContent } from '#content/server'

export default defineEventHandler(async (event) => {
  const docs = await serverQueryContent(event).find()
  const baseUrl = 'https://kerjakantor.my.id'

  const sitemapEntries = docs
    .filter((doc) => doc._path && !doc._path.startsWith('/_'))
    .map((doc) => {
      const url = `${baseUrl}${doc._path}`
      const lastmod = doc.date ? new Date(doc.date).toISOString().split('T')[0] : new Date().toISOString().split('T')[0]
      return `  <url>\n    <loc>${url}</loc>\n    <lastmod>${lastmod}</lastmod>\n    <changefreq>weekly</changefreq>\n    <priority>0.8</priority>\n  </url>`
    })

  const staticPages = [
    '',
    '/kebijakan-privasi',
    '/ketentuan-layanan',
    '/tentang-kami',
    '/kontak',
    '/kategori/excel',
    '/kategori/google-sheets'
  ].map((path) => {
    return `  <url>\n    <loc>${baseUrl}${path}</loc>\n    <lastmod>${new Date().toISOString().split('T')[0]}</lastmod>\n    <changefreq>daily</changefreq>\n    <priority>${path === '' ? '1.0' : '0.5'}</priority>\n  </url>`
  })

  const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${staticPages.join('\n')}
${sitemapEntries.join('\n')}
</urlset>`

  setHeader(event, 'content-type', 'application/xml')
  return sitemap
})
