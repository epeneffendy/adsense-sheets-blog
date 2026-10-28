<template>
  <div class="space-y-12">
    <!-- 1. Hero Feature Section (Editorial Portal Media Style: 2 Kolom) -->
    <section v-if="featuredArticle" class="grid grid-cols-1 lg:grid-cols-12 gap-8 items-stretch">
      <!-- Main Featured Lead Card (8 Kolom) -->
      <article class="lg:col-span-8 bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col justify-between group hover:border-emerald-600 transition duration-300">
        <div class="p-8 sm:p-12 space-y-6">
          <div class="flex items-center gap-3 text-xs">
            <span class="px-3 py-1 rounded-full font-bold uppercase tracking-wider text-[11px] bg-emerald-600 text-white shadow-xs">
              Tajuk Utama
            </span>
            <span class="font-semibold text-emerald-800 bg-emerald-50 px-2.5 py-0.5 rounded-md border border-emerald-200">
              {{ featuredArticle.category }}
            </span>
            <span class="text-slate-400">•</span>
            <span class="text-slate-500">{{ featuredArticle.readTime || '8 menit baca' }}</span>
          </div>

          <h1 class="text-2xl sm:text-4xl font-black text-slate-900 group-hover:text-emerald-700 transition leading-tight tracking-tight">
            <NuxtLink :to="featuredArticle._path">
              {{ featuredArticle.title }}
            </NuxtLink>
          </h1>

          <p class="text-slate-600 text-sm sm:text-base leading-relaxed line-clamp-3">
            {{ featuredArticle.description }}
          </p>
        </div>

        <div class="px-8 sm:px-12 py-6 bg-slate-50/80 border-t border-slate-100 flex items-center justify-between">
          <div class="flex items-center gap-3">
            <div class="w-9 h-9 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center text-xs">
              RP
            </div>
            <div>
              <div class="text-xs font-bold text-slate-800">Dewan Redaksi KerjaKantor</div>
              <div class="text-[11px] text-slate-400">Ditinjau &amp; Diuji pada Spreadsheet</div>
            </div>
          </div>

          <NuxtLink 
            :to="featuredArticle._path"
            class="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-sm transition"
          >
            Baca Selengkapnya &rarr;
          </NuxtLink>
        </div>
      </article>

      <!-- Sidebar Editorial Feed (4 Kolom: "TAJUK TERKINI") -->
      <aside class="lg:col-span-4 bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 flex flex-col justify-between shadow-sm">
        <div>
          <div class="flex items-center justify-between border-b border-slate-100 pb-4 mb-5">
            <div class="flex items-center gap-2">
              <span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
              <h2 class="text-xs font-black uppercase tracking-wider text-slate-900">TAJUK TERKINI</h2>
            </div>
            <span class="text-[11px] text-slate-400">Update Harian</span>
          </div>

          <div class="divide-y divide-slate-100">
            <article 
              v-for="item in sidebarArticles" 
              :key="item._path"
              class="py-4 first:pt-0 last:pb-0 group"
            >
              <div class="flex items-center gap-2 text-[11px] text-slate-400 mb-1">
                <span class="font-bold text-emerald-700">{{ item.category }}</span>
                <span>•</span>
                <span>{{ item.readTime || '5 menit' }}</span>
              </div>
              <h3 class="text-sm font-bold text-slate-900 group-hover:text-emerald-700 transition leading-snug line-clamp-2">
                <NuxtLink :to="item._path">
                  {{ item.title }}
                </NuxtLink>
              </h3>
            </article>
          </div>
        </div>

        <div class="pt-6 border-t border-slate-100 mt-6">
          <NuxtLink 
            to="/kategori/excel"
            class="block w-full py-2.5 text-center text-xs font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 rounded-xl transition"
          >
            Jelajahi Rubrik Excel &rarr;
          </NuxtLink>
        </div>
      </aside>
    </section>

    <!-- 2. Grid Koleksi Artikel Utama + Pagination -->
    <section class="space-y-8">
      <!-- Section Header -->
      <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h2 class="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2.5">
            <span>Lembar Panduan &amp; Dokumentasi</span>
          </h2>
          <p class="text-xs text-slate-500 mt-1">
            Menampilkan halaman {{ currentPage }} dari {{ totalPages }} (Total {{ allArticles?.length || 0 }} panduan)
          </p>
        </div>

        <!-- Filter Shortcut Pills -->
        <div class="flex items-center gap-2 text-xs font-semibold">
          <span class="text-slate-400 mr-1 hidden sm:inline">Kategori:</span>
          <NuxtLink to="/kategori/excel" class="px-3 py-1.5 rounded-lg bg-white border border-slate-200 hover:border-emerald-500 hover:text-emerald-700 transition">
            Excel (14)
          </NuxtLink>
          <NuxtLink to="/kategori/google-sheets" class="px-3 py-1.5 rounded-lg bg-white border border-slate-200 hover:border-emerald-500 hover:text-emerald-700 transition">
            Google Sheets (12)
          </NuxtLink>
        </div>
      </div>

      <!-- Loading State -->
      <div v-if="pending" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        <div v-for="i in 6" :key="i" class="p-6 bg-white rounded-2xl border border-slate-200 animate-pulse space-y-4">
          <div class="h-4 bg-slate-200 rounded w-1/4"></div>
          <div class="h-6 bg-slate-200 rounded w-3/4"></div>
          <div class="h-16 bg-slate-100 rounded w-full"></div>
        </div>
      </div>

      <!-- 3-Column Modern Grid Cards -->
      <div v-else class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        <article
          v-for="article in paginatedArticles"
          :key="article._path"
          class="flex flex-col bg-white rounded-2xl border border-slate-200 hover:border-emerald-500 hover:shadow-lg transition duration-200 group overflow-hidden"
        >
          <div class="p-6 flex-1 flex flex-col justify-between space-y-4">
            <div class="space-y-3">
              <div class="flex items-center justify-between text-xs">
                <span
                  :class="[
                    'px-2.5 py-0.5 rounded-md font-bold text-[11px] tracking-wide',
                    article.category === 'Excel' 
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' 
                      : 'bg-teal-50 text-teal-800 border border-teal-200'
                  ]"
                >
                  {{ article.category }}
                </span>
                <span class="text-slate-400 text-[11px]">{{ article.readTime || '6 menit' }}</span>
              </div>

              <h3 class="text-base font-bold text-slate-900 group-hover:text-emerald-700 transition line-clamp-2 leading-snug">
                <NuxtLink :to="article._path">
                  {{ article.title }}
                </NuxtLink>
              </h3>

              <p class="text-xs text-slate-600 line-clamp-3 leading-relaxed">
                {{ article.description }}
              </p>
            </div>

            <div class="pt-4 border-t border-slate-100 flex items-center justify-between text-xs">
              <time class="text-slate-400 text-[11px]">
                {{ article.date ? new Date(article.date).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : 'Terbaru' }}
              </time>

              <NuxtLink 
                :to="article._path"
                class="font-bold text-emerald-600 hover:text-emerald-800 group-hover:translate-x-0.5 transition-transform"
              >
                Baca &rarr;
              </NuxtLink>
            </div>
          </div>
        </article>
      </div>

      <!-- 3. Pagination Controls (AdSense Compliant) -->
      <nav v-if="totalPages > 1" class="pt-8 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-4 select-none">
        <div class="text-xs text-slate-500">
          Menampilkan baris {{ (currentPage - 1) * perPage + 1 }} – {{ Math.min(currentPage * perPage, allArticles?.length || 0) }} dari total {{ allArticles?.length || 0 }} artikel
        </div>

        <div class="flex items-center gap-1.5">
          <!-- Tombol Prev -->
          <button
            @click="goToPage(currentPage - 1)"
            :disabled="currentPage === 1"
            class="px-3 py-2 rounded-xl text-xs font-bold border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition"
          >
            &larr; Sebelumnya
          </button>

          <!-- Nomor Halaman -->
          <button
            v-for="p in totalPages"
            :key="p"
            @click="goToPage(p)"
            :class="[
              'w-9 h-9 rounded-xl text-xs font-bold transition flex items-center justify-center',
              currentPage === p
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-50'
            ]"
          >
            {{ p }}
          </button>

          <!-- Tombol Next -->
          <button
            @click="goToPage(currentPage + 1)"
            :disabled="currentPage === totalPages"
            class="px-3 py-2 rounded-xl text-xs font-bold border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed transition"
          >
            Selanjutnya &rarr;
          </button>
        </div>
      </nav>
    </section>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'

const { data: allArticles, pending } = await useAsyncData('articles-index', () => {
  return queryContent('articles')
    .sort({ date: -1 })
    .find()
})

// Artikel Unggulan (Hero Lead)
const featuredArticle = computed(() => {
  return allArticles.value?.[0] || null
})

// Artikel Tajuk Terkini Sidebar (3 artikel berikutnya)
const sidebarArticles = computed(() => {
  return allArticles.value?.slice(1, 4) || []
})

// Pagination Logic
const currentPage = ref(1)
const perPage = 9 // 9 card per halaman (kelipatan 3 kolom pas)

const totalPages = computed(() => {
  if (!allArticles.value) return 1
  return Math.ceil(allArticles.value.length / perPage)
})

const paginatedArticles = computed(() => {
  if (!allArticles.value) return []
  const start = (currentPage.value - 1) * perPage
  return allArticles.value.slice(start, start + perPage)
})

const goToPage = (page: number) => {
  if (page >= 1 && page <= totalPages.value) {
    currentPage.value = page
    window.scrollTo({ top: 400, behavior: 'smooth' })
  }
}

useHead({
  title: 'KerjaKantor - Portal Dokumentasi Excel & Google Sheets Terlengkap',
  meta: [
    { name: 'description', content: 'Publikasi edukasi spreadsheet profesional: tutorial formula Excel, Google Sheets, dan otomatisasi data kantor.' }
  ]
})
</script>
