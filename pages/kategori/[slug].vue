<template>
  <div class="space-y-8">
    <!-- Header Kategori -->
    <div class="bg-white p-8 rounded-3xl border border-slate-200 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
      <div>
        <div class="flex items-center gap-2 text-xs font-bold text-emerald-700 uppercase tracking-wider mb-2">
          <span>RUBRIK KHUSUS</span>
          <span>•</span>
          <span>{{ filteredArticles.length }} DOKUMENTASI</span>
        </div>
        <h1 class="text-3xl font-black text-slate-900 tracking-tight">{{ categoryName }}</h1>
        <p class="text-xs text-slate-500 mt-1 max-w-xl">
          Kumpulan panduan terstruktur, studi kasus formula, dan teknik otomasi khusus ekosistem {{ categoryName }}.
        </p>
      </div>

      <NuxtLink 
        to="/" 
        class="self-start md:self-auto px-4 py-2 rounded-xl bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-xs font-bold text-slate-700 transition"
      >
        &larr; Lihat Semua Kategori
      </NuxtLink>
    </div>

    <!-- Empty State -->
    <div v-if="!filteredArticles || filteredArticles.length === 0" class="p-16 text-center bg-white rounded-2xl border border-slate-200">
      <p class="text-slate-500 text-sm">Belum ada artikel dalam kategori ini.</p>
    </div>

    <!-- 3-Column Grid + Pagination -->
    <div v-else class="space-y-8">
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        <article
          v-for="article in paginatedArticles"
          :key="article._path"
          class="flex flex-col bg-white rounded-2xl border border-slate-200 hover:border-emerald-500 hover:shadow-lg transition duration-200 group overflow-hidden"
        >
          <div class="p-6 flex-1 flex flex-col justify-between space-y-4">
            <div class="space-y-3">
              <div class="flex items-center justify-between text-xs">
                <span class="text-slate-400 text-[11px] font-medium">
                  {{ article.date ? new Date(article.date).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : 'Terbaru' }}
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

            <div class="pt-4 border-t border-slate-100 flex items-center justify-end text-xs">
              <NuxtLink 
                :to="article._path"
                class="font-bold text-emerald-600 hover:text-emerald-800 group-hover:translate-x-0.5 transition-transform"
              >
                Baca Panduan &rarr;
              </NuxtLink>
            </div>
          </div>
        </article>
      </div>

      <!-- Pagination Kategori -->
      <nav v-if="totalPages > 1" class="pt-6 border-t border-slate-200 flex items-center justify-between text-xs">
        <div class="text-slate-500">
          Halaman {{ currentPage }} dari {{ totalPages }}
        </div>

        <div class="flex items-center gap-1.5">
          <button
            @click="currentPage--"
            :disabled="currentPage === 1"
            class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 transition font-bold"
          >
            &larr; Prev
          </button>
          <button
            v-for="p in totalPages"
            :key="p"
            @click="currentPage = p"
            :class="[
              'w-8 h-8 rounded-lg font-bold flex items-center justify-center transition',
              currentPage === p ? 'bg-emerald-600 text-white' : 'bg-white border border-slate-200 text-slate-700'
            ]"
          >
            {{ p }}
          </button>
          <button
            @click="currentPage++"
            :disabled="currentPage === totalPages"
            class="px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 disabled:opacity-40 transition font-bold"
          >
            Next &rarr;
          </button>
        </div>
      </nav>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'

const route = useRoute()
const slug = computed(() => (route.params.slug as string) || '')
const categoryName = computed(() => {
  if (slug.value === 'google-sheets') return 'Google Sheets'
  if (slug.value === 'excel') return 'Microsoft Excel'
  return slug.value
})

const { data: allArticles } = await useAsyncData(`cat-${slug.value}`, () => {
  return queryContent('articles').sort({ date: -1 }).find()
})

const filteredArticles = computed(() => {
  if (!allArticles.value) return []
  return allArticles.value.filter(a => {
    const cat = (a.category || '').toLowerCase().replace(/\s+/g, '-')
    return cat === slug.value.toLowerCase()
  })
})

const currentPage = ref(1)
const perPage = 9

const totalPages = computed(() => {
  if (!filteredArticles.value.length) return 1
  return Math.ceil(filteredArticles.value.length / perPage)
})

const paginatedArticles = computed(() => {
  const start = (currentPage.value - 1) * perPage
  return filteredArticles.value.slice(start, start + perPage)
})

useHead({
  title: `Panduan ${categoryName.value} - KerjaKantor`,
  meta: [
    { name: 'description', content: `Koleksi panduan teknis dan tutorial formula ${categoryName.value} terlengkap.` }
  ]
})
</script>
