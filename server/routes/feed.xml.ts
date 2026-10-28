import { serverQueryContent } from '#content/server'

export default defineEventHandler(async (event) => {
  const docs = await serverQueryContent(event)
    .sort({ date: -1 })
    .find()

  const baseUrl = 'https://kerjakantor.my.id'

  const items = docs
    .filter((doc) => doc._path && !doc._path.startsWith('/_'))
    .map((doc) => {
      return `    <item>
      <title><![CDATA[${doc.title}]]></title>
      <link>${baseUrl}${doc._path}</link>
      <guid>${baseUrl}${doc._path}</guid>
      <pubDate>${doc.date ? new Date(doc.date).toUTCString() : new Date().toUTCString()}</pubDate>
      <description><![CDATA[${doc.description || ''}]]></description>
    </item>`
    })
    .join('\n')

  const feed = `<?xml version="1.0" encoding="UTF-8" ?>
<rss version="2.0">
  <channel>
    <title>KerjaKantor - Panduan Excel &amp; Google Sheets</title>
    <link>${baseUrl}</link>
    <description>Tutorial formula Excel, Google Sheets, dan otomasi spreadsheet terlengkap.</description>
    <language>id</language>
    <lastBuildDate>${new Date().toUTCString()}</lastBuildDate>
${items}
  </channel>
</rss>`

  setHeader(event, 'content-type', 'application/xml')
  return feed
})
