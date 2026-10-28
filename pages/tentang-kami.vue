<template>
  <div class="max-w-3xl mx-auto bg-white p-8 md:p-12 rounded-2xl border border-slate-200 shadow-sm space-y-6">
    <h1 class="text-3xl font-extrabold text-slate-900 tracking-tight">Tentang Kami</h1>

    <div class="space-y-4 text-sm text-slate-700 leading-relaxed">
      <p>
        <strong>KerjaKantor</strong> adalah portal edukasi teknologi yang berfokus pada pemanfaatan spreadsheet (Microsoft Excel dan Google Sheets) untuk meningkatkan efisiensi dan otomasi kerja sehari-hari.
      </p>

      <h2 class="text-lg font-bold text-slate-900 pt-4">Misi Kami</h2>
      <p>
        Banyak waktu kerja kantor terbuang karena pengolahan data manual yang repetitif. Misi kami adalah menyederhanakan konsep formula lanjutan (seperti XLOOKUP, INDEX-MATCH ganda, ARRAYFORMULA, hingga skrip otomasi Google Apps Script) ke dalam panduan praktis berbahasa Indonesia yang mudah diterapkan oleh staf operasional, akuntan, maupun analis data.
      </p>

      <h2 class="text-lg font-bold text-slate-900 pt-4">Standar Konten</h2>
      <p>
        Setiap artikel di KerjaKantor disusun berdasarkan studi kasus nyata di dunia kerja, dilengkapi penjelasan logika fungsi langkah demi langkah, dan telah diverifikasi langsung pada versi Excel maupun Google Sheets terbaru.
      </p>

      <div class="p-6 bg-slate-50 border border-slate-200 rounded-xl mt-6">
        <h3 class="font-bold text-slate-900 mb-2">Punya Pertanyaan atau Ide Tutorial?</h3>
        <p class="text-xs text-slate-600">
          Kami menyambut masukan, pertanyaan teknis, dan kerja sama konten melalui formulir di halaman <NuxtLink to="/kontak" class="text-emerald-600 font-semibold underline">Kontak Kami</NuxtLink>.
        </p>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
useHead({
  title: 'Tentang Kami - KerjaKantor',
  meta: [
    { name: 'description', content: 'Tentang portal edukasi formula spreadsheet KerjaKantor.' }
  ]
})
</script>
