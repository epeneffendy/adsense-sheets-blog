<template>
  <div class="max-w-3xl mx-auto bg-white p-8 md:p-12 rounded-2xl border border-slate-200 shadow-sm space-y-6">
    <h1 class="text-3xl font-extrabold text-slate-900 tracking-tight">Hubungi Kami</h1>
    <p class="text-sm text-slate-600">
      Silakan hubungi tim redaksi KerjaKantor jika Anda memiliki pertanyaan teknis seputar tutorial, kritik, saran, atau peluang kolaborasi.
    </p>

    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4">
      <div class="p-6 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
        <div class="text-xs font-semibold text-emerald-700 uppercase tracking-wider">Email Redaksi</div>
        <div class="text-slate-900 font-medium">redaksi@kerjakantor.my.id</div>
        <p class="text-xs text-slate-500">Respon dalam 1-2 hari kerja.</p>
      </div>

      <div class="p-6 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
        <div class="text-xs font-semibold text-emerald-700 uppercase tracking-wider">Lokasi</div>
        <div class="text-slate-900 font-medium">Jakarta / Remote, Indonesia</div>
        <p class="text-xs text-slate-500">Operasional Senin - Jumat.</p>
      </div>
    </div>

    <!-- Static Mock Form -->
    <div class="pt-6 border-t border-slate-100 space-y-4">
      <h2 class="text-lg font-bold text-slate-900">Kirim Pesan Cepat</h2>
      <div class="space-y-4">
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">Nama Lengkap</label>
          <input type="text" placeholder="Nama Anda" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600" />
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">Alamat Email</label>
          <input type="email" placeholder="nama@email.com" class="w-full px-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600" />
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">Pesan</label>
          <textarea rows="4" placeholder="Tuliskan pesan atau topik formula yang ingin dibahas..." class="w-full px-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"></textarea>
        </div>
        <button type="button" class="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-sm rounded-lg shadow-sm transition">
          Kirim Pesan
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
useHead({
  title: 'Kontak Kami - KerjaKantor',
  meta: [
    { name: 'description', content: 'Hubungi tim redaksi KerjaKantor untuk pertanyaan dan kolaborasi spreadsheet.' }
  ]
})
</script>
