<script setup>
import { ref, computed } from 'vue'
import { exportReport } from './report/generateReport'

const pipeDiameter = ref('')
const deviceNumber = ref('')
const measurementsCount = ref(3)

const isValid = computed(() => {
  const countOk = Number(measurementsCount.value) >= 1
  return Boolean(deviceNumber.value.trim()) && countOk
})

function onExport() {
  exportReport({
    pipeDiameter: pipeDiameter.value,
    deviceNumber: deviceNumber.value.trim(),
    measurementsCount: Number(measurementsCount.value) || 1,
  })
}
</script>

<template>
  <div class="container">
    <h1>Калькулятор</h1>

    <form class="form" @submit.prevent="onExport">
      <label class="field">
        <span>Диаметр трубы</span>
        <input v-model="pipeDiameter" type="text" placeholder="например, 50 мм" />
      </label>

      <label class="field">
        <span>Номер устройства</span>
        <input v-model="deviceNumber" type="text" placeholder="Заводской номер" required />
      </label>

      <label class="field">
        <span>Количество измерений</span>
        <input v-model.number="measurementsCount" type="number" min="1" step="1" />
      </label>

      <button type="submit" :disabled="!isValid">Выгрузить отчёт в формате xlsx</button>
    </form>
  </div>
</template>

<style scoped>
.container {
  max-width: 520px;
  margin: 40px auto;
  padding: 24px;
  border: 1px solid #e5e7eb;
  border-radius: 12px;
  background: #ffffff;
  box-shadow: 0 8px 20px rgba(0,0,0,0.04);
}

h1 {
  margin: 0 0 16px;
  font-size: 24px;
}

.form {
  display: grid;
  gap: 16px;
}

.field {
  display: grid;
  gap: 6px;
}

input[type="text"], input[type="number"] {
  padding: 10px 12px;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  font-size: 14px;
}

button {
  padding: 12px 16px;
  background: #4f46e5;
  color: white;
  border: none;
  border-radius: 10px;
  cursor: pointer;
}

button:disabled {
  background: #9ca3af;
  cursor: not-allowed;
}
</style>
