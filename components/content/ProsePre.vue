<template>
  <div class="code-block-wrapper relative my-6 rounded-2xl overflow-hidden border border-slate-800 bg-slate-950 shadow-xl group">
    <!-- Header Mac-style terminal -->
    <div class="flex items-center justify-between px-4 py-3 bg-slate-900 border-b border-slate-800/80 text-xs text-slate-300 select-none">
      <div class="flex items-center gap-2.5">
        <div class="flex gap-1.5 items-center">
          <span class="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block"></span>
          <span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>
          <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>
        </div>
        <span class="ml-2 font-mono text-[11px] font-bold tracking-wider text-slate-200 uppercase">
          {{ displayLanguage }}
        </span>
      </div>

      <!-- Copy Button -->
      <button
        @click="copyCode"
        type="button"
        class="inline-flex items-center gap-1.5 px-3 py-1 rounded-lg bg-slate-800 hover:bg-emerald-600 text-slate-200 hover:text-white transition-all duration-200 font-sans text-[11px] font-semibold cursor-pointer shadow-xs active:scale-95"
      >
        <span v-if="copied" class="text-emerald-300 font-bold flex items-center gap-1">
          ✓ Tersalin
        </span>
        <span v-else class="flex items-center gap-1">
          <svg class="w-3.5 h-3.5 text-slate-400 group-hover:text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
          </svg>
          Salin Rumus
        </span>
      </button>
    </div>

    <!-- Code Block Content -->
    <div 
      ref="codeContainer"
      class="code-container overflow-x-auto p-5 text-emerald-300 font-mono text-[14px] leading-relaxed tracking-wide"
    >
      <slot />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'

const props = defineProps({
  code: { type: String, default: '' },
  language: { type: String, default: '' },
  filename: { type: String, default: '' },
  highlights: { type: Array as () => number[], default: () => [] },
  meta: { type: String, default: '' }
})

const copied = ref(false)
const codeContainer = ref<HTMLElement | null>(null)

const displayLanguage = computed(() => {
  if (props.language) {
    if (props.language === 'excel') return 'EXCEL / SPREADSHEET'
    return props.language
  }
  return props.filename || 'FORMULA'
})

const copyCode = async () => {
  const text = props.code || codeContainer.value?.innerText || ''
  if (!text) return

  try {
    await navigator.clipboard.writeText(text.trim())
    copied.value = true
    setTimeout(() => {
      copied.value = false
    }, 2000)
  } catch (err) {
    console.error('Gagal menyalin:', err)
  }
}
</script>

<style>
/* Memaksa kontras warna teks di dalam blok kode agar selalu terang dan mudah dibaca */
.code-container pre,
.code-container code {
  background: transparent !important;
  color: #34d399 !important; /* emerald-400 terang */
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace !important;
  font-size: 14px !important;
  line-height: 1.7 !important;
}

.code-container span {
  color: #a7f3d0 !important; /* emerald-200 sangat terang untuk teks formula */
}

/* Pewarnaan komentar rumus */
.code-container span:has(span) {
  color: #a7f3d0 !important;
}

/* Jika ada Shiki span token, paksa fallback kontras tinggi */
.code-container .shiki span {
  color: #6ee7b7 !important; /* emerald-300 */
}

/* Komentar baris (// ...) dibuat abu-abu terang */
.code-container .line:has(span:first-child) {
  display: block;
}
</style>
