<template>
  <div class="min-h-screen bg-[#F8FAFC] text-slate-800 flex flex-col font-sans antialiased selection:bg-emerald-600 selection:text-white">
    <!-- Top Utility Bar (Portal Media Style) -->
    <div class="bg-slate-900 text-slate-400 text-xs py-2 px-4 border-b border-slate-800">
      <div class="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
        <div class="flex items-center gap-4 text-[11px]">
          <span class="inline-flex items-center gap-1.5 font-medium text-emerald-400">
            <span class="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
            PORTAL DOKUMENTASI SPREADSHEET
          </span>
          <span class="hidden md:inline text-slate-600">•</span>
          <span class="hidden md:inline text-slate-400">Edisi Praktik Microsoft Excel &amp; Google Sheets</span>
        </div>
        <div class="flex items-center gap-4 text-[11px] font-medium">
          <span>{{ todayFormatted }}</span>
          <span class="text-slate-700">|</span>
          <NuxtLink to="/tentang-kami" class="hover:text-emerald-400 transition">Standar Redaksi</NuxtLink>
        </div>
      </div>
    </div>

    <!-- Main Navigation Header -->
    <header class="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-slate-200 shadow-xs">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 h-20 flex items-center justify-between">
        <!-- Brand Logo Portal -->
        <NuxtLink to="/" class="flex items-center gap-3 group">
          <div class="w-11 h-11 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-800 flex items-center justify-center text-white font-black text-xl shadow-md shadow-emerald-900/10 group-hover:scale-105 transition-transform">
            fx
          </div>
          <div class="flex flex-col">
            <span class="font-black text-2xl tracking-tight text-slate-900 group-hover:text-emerald-700 transition leading-none">
              RUMUS<span class="text-emerald-600">PRO</span>
            </span>
            <span class="text-[10px] uppercase font-bold tracking-widest text-slate-600 mt-1">
              Jurnal Spreadsheet &amp; Otomasi Data
            </span>
          </div>
        </NuxtLink>

        <!-- Category Menu Links -->
        <nav class="hidden md:flex items-center gap-1 text-sm font-bold text-slate-700">
          <NuxtLink 
            to="/" 
            class="px-3.5 py-2 rounded-lg hover:text-emerald-700 hover:bg-slate-50 transition"
            exact-active-class="text-emerald-700 bg-emerald-50/60"
          >
            Beranda
          </NuxtLink>
          <NuxtLink 
            to="/kategori/excel" 
            class="px-3.5 py-2 rounded-lg hover:text-emerald-700 hover:bg-slate-50 transition"
            active-class="text-emerald-700 bg-emerald-50/60"
          >
            Microsoft Excel
          </NuxtLink>
          <NuxtLink 
            to="/kategori/google-sheets" 
            class="px-3.5 py-2 rounded-lg hover:text-emerald-700 hover:bg-slate-50 transition"
            active-class="text-emerald-700 bg-emerald-50/60"
          >
            Google Sheets
          </NuxtLink>
          <NuxtLink 
            to="/tentang-kami" 
            class="px-3.5 py-2 rounded-lg hover:text-emerald-700 hover:bg-slate-50 transition"
            active-class="text-emerald-700 bg-emerald-50/60"
          >
            Tentang Kami
          </NuxtLink>
        </nav>

        <!-- Search / Action Badge -->
        <div class="flex items-center gap-3">
          <span class="hidden sm:inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-emerald-100/70 text-emerald-800 border border-emerald-200">
            27 Panduan Terverifikasi
          </span>
        </div>
      </div>
    </header>

    <!-- Main Workspace -->
    <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 py-8">
      <slot />
    </main>

    <!-- Footer AdSense Portal Style -->
    <footer class="bg-slate-900 text-slate-400 border-t border-slate-800 mt-20 text-sm">
      <div class="max-w-7xl mx-auto px-4 sm:px-6 py-14 space-y-10">
        <div class="grid grid-cols-1 md:grid-cols-4 gap-8">
          <!-- Col 1: About Portal -->
          <div class="space-y-4 md:col-span-2">
            <div class="flex items-center gap-3">
              <div class="w-8 h-8 rounded-lg bg-emerald-600 text-white font-bold flex items-center justify-center text-xs">
                fx
              </div>
              <span class="font-extrabold text-xl text-white tracking-tight">KERJAKANTOR</span>
            </div>
            <p class="text-xs text-slate-400 leading-relaxed max-w-md">
              KerjaKantor adalah publikasi edukasi teknologi spreadsheet terstruktur yang menyajikan panduan mendalam untuk formula Microsoft Excel, Google Sheets, dan Google Apps Script bagi profesional, akuntan, dan analis data.
            </p>
          </div>

          <!-- Col 2: Kategori Utama -->
          <div class="space-y-3">
            <h4 class="text-xs font-bold uppercase tracking-wider text-white">Rubrik Panduan</h4>
            <ul class="space-y-2 text-xs">
              <li><NuxtLink to="/kategori/excel" class="hover:text-emerald-400 transition">Tutorial Microsoft Excel</NuxtLink></li>
              <li><NuxtLink to="/kategori/google-sheets" class="hover:text-emerald-400 transition">Tutorial Google Sheets</NuxtLink></li>
              <li><NuxtLink to="/" class="hover:text-emerald-400 transition">Formula Index &amp; Match</NuxtLink></li>
              <li><NuxtLink to="/" class="hover:text-emerald-400 transition">Otomasi Apps Script</NuxtLink></li>
            </ul>
          </div>

          <!-- Col 3: Kebijakan & Transparansi (Wajib AdSense) -->
          <div class="space-y-3">
            <h4 class="text-xs font-bold uppercase tracking-wider text-white">Standar &amp; Hukum</h4>
            <ul class="space-y-2 text-xs">
              <li><NuxtLink to="/kebijakan-privasi" class="hover:text-emerald-400 transition">Kebijakan Privasi</NuxtLink></li>
              <li><NuxtLink to="/ketentuan-layanan" class="hover:text-emerald-400 transition">Ketentuan Layanan</NuxtLink></li>
              <li><NuxtLink to="/tentang-kami" class="hover:text-emerald-400 transition">Tentang Kami</NuxtLink></li>
              <li><NuxtLink to="/kontak" class="hover:text-emerald-400 transition">Kontak Redaksi</NuxtLink></li>
            </ul>
          </div>
        </div>

        <div class="pt-8 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between text-xs text-slate-400 gap-4">
          <p>&copy; {{ currentYear }} KerjaKantor. Seluruh materi ditinjau untuk tujuan edukasi operasional.</p>
          <div class="flex items-center gap-4 text-xs">
            <NuxtLink to="/sitemap.xml" class="hover:text-emerald-400">Peta Situs (Sitemap)</NuxtLink>
            <span>•</span>
            <NuxtLink to="/feed.xml" class="hover:text-emerald-400">RSS Feed</NuxtLink>
          </div>
        </div>
      </div>
    </footer>
  </div>
</template>

<script setup lang="ts">
const currentYear = new Date().getFullYear()
const todayFormatted = new Date().toLocaleDateString('id-ID', {
  weekday: 'long',
  year: 'numeric',
  month: 'long',
  day: 'numeric'
})
</script>
