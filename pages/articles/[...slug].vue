<template>
  <main class="max-w-4xl mx-auto py-4">
    <!-- Breadcrumbs -->
    <nav class="flex items-center gap-2 text-xs text-slate-500 mb-6">
      <NuxtLink to="/" class="hover:text-emerald-600 transition">Beranda</NuxtLink>
      <span>/</span>
      <NuxtLink 
        v-if="doc?.category" 
        :to="`/kategori/${doc.category.toLowerCase().replace(/\s+/g, '-')}`"
        class="hover:text-emerald-600 transition capitalize"
      >
        {{ doc.category }}
      </NuxtLink>
      <span v-else class="text-slate-700 capitalize">Artikel</span>
    </nav>

    <!-- Content Document Viewer -->
    <ContentDoc v-slot="{ doc }">
      <article class="bg-white p-6 sm:p-12 rounded-3xl border border-slate-200/90 shadow-sm space-y-8">
        <!-- Header -->
        <header class="space-y-4 border-b border-slate-100 pb-8">
          <div class="flex flex-wrap items-center gap-3 text-xs text-slate-500">
            <span 
              :class="[
                'px-2.5 py-1 rounded-md font-semibold text-[11px] tracking-wide',
                doc.category === 'Excel' 
                  ? 'bg-emerald-50 text-emerald-700 border border-emerald-200/60' 
                  : 'bg-teal-50 text-teal-700 border border-teal-200/60'
              ]"
            >
              {{ doc.category || 'Spreadsheet' }}
            </span>
            <span>•</span>
            <time>{{ doc.date ? new Date(doc.date).toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric' }) : '' }}</time>
            <span>•</span>
            <span>{{ doc.readTime || '6 menit baca' }}</span>
          </div>

          <h1 class="text-3xl sm:text-4xl font-black text-slate-900 tracking-tight leading-tight">
            {{ doc.title }}
          </h1>

          <p class="text-base text-slate-600 leading-relaxed font-normal">
            {{ doc.description }}
          </p>
        </header>

        <!-- Body Render Markdown with Clean Typography -->
        <div class="prose prose-slate max-w-none 
                    prose-headings:font-black prose-headings:tracking-tight prose-headings:text-slate-900
                    prose-h2:text-2xl prose-h2:mt-10 prose-h2:mb-4 prose-h2:pb-2 prose-h2:border-b prose-h2:border-slate-100
                    prose-h3:text-xl prose-h3:mt-6 prose-h3:mb-3
                    prose-p:text-slate-700 prose-p:leading-relaxed prose-p:text-[15px]
                    prose-ul:my-4 prose-li:text-slate-700 prose-li:my-1
                    prose-table:w-full prose-table:border prose-table:border-slate-200 prose-table:rounded-xl prose-table:overflow-hidden
                    prose-th:bg-slate-100 prose-th:p-3 prose-th:text-xs prose-th:font-bold prose-th:text-slate-700 prose-th:border-b
                    prose-td:p-3 prose-td:text-xs prose-td:border-b prose-td:border-slate-100">
          <ContentRenderer :value="doc" />
        </div>

        <!-- Author / Footer Box -->
        <footer class="mt-14 pt-8 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
          <div class="flex items-center gap-3">
            <div class="w-8 h-8 rounded-full bg-emerald-100 text-emerald-700 font-bold flex items-center justify-center">
              RP
            </div>
            <div>
              <div class="font-semibold text-slate-800">Tim Edukasi KerjaKantor</div>
              <div class="text-[11px] text-slate-400">Ditinjau untuk Excel & Google Sheets</div>
            </div>
          </div>

          <NuxtLink 
            to="/" 
            class="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 font-semibold text-slate-700 transition"
          >
            &larr; Kembali ke daftar panduan
          </NuxtLink>
        </footer>
      </article>
    </ContentDoc>
  </main>
</template>

<script setup lang="ts">
const route = useRoute()
const { data: doc } = await useAsyncData(`content-${route.path}`, () => {
  return queryContent(route.path).findOne()
})

useHead({
  title: doc.value ? `${doc.value.title} - KerjaKantor` : 'Panduan Spreadsheet - KerjaKantor',
  meta: [
    { name: 'description', content: doc.value?.description || 'Panduan praktis formula spreadsheet.' }
  ]
})
</script>

<style>
/* Pastikan tag pre dan code di dalam prose tidak menimpa background dark */
.prose pre {
  background-color: transparent !important;
  padding: 0 !important;
  margin: 0 !important;
}

/* Pastikan warna teks dalam blok kode terlihat terang dan tajam */
.prose pre code span {
  color: #a7f3d0 !important; /* Hijau muda terang kontras tinggi */
}
</style>
