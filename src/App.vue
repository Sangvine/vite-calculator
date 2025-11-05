<script setup>
import { ref, computed } from 'vue'
import { exportReport, DIAMETER_OPTIONS } from './report/generateReport'

const pipeDiameter = ref('50 мм') // значение по умолчанию
const deviceNumber = ref('')
const measurementsCount90 = ref(3)
const measurementsCount50 = ref(3)
const measurementsCount10 = ref(3)
const measurementsCount2 = ref(3)

const isValid = computed(() => {
  const count90 = Number(measurementsCount90.value) || 0
  const count50 = Number(measurementsCount50.value) || 0
  const count10 = Number(measurementsCount10.value) || 0
  const count2 = Number(measurementsCount2.value) || 0
  // Все значения должны быть >= 0
  const allNonNegative = count90 >= 0 && count50 >= 0 && count10 >= 0 && count2 >= 0
  return Boolean(deviceNumber.value.trim()) && allNonNegative && Boolean(pipeDiameter.value)
})

function onExport() {
  exportReport({
    pipeDiameter: pipeDiameter.value,
    deviceNumber: deviceNumber.value.trim(),
    measurementsCount: {
      flow90: Number(measurementsCount90.value) || 0,
      flow50: Number(measurementsCount50.value) || 0,
      flow10: Number(measurementsCount10.value) || 0,
      flow2: Number(measurementsCount2.value) || 0,
    },
  })
}
</script>

<template>
  <div class="container">
    <h1>Калькулятор</h1>

    <form class="form" @submit.prevent="onExport">
      <label class="field">
        <span>Диаметр трубы</span>
        <select v-model="pipeDiameter">
          <option value="">Выберите диаметр</option>
          <option v-for="option in DIAMETER_OPTIONS" :key="option.value" :value="option.value">
            {{ option.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Номер устройства</span>
        <input v-model="deviceNumber" type="text" placeholder="Заводской номер" required />
      </label>

      <label class="field">
        <span>Количество измерений (90%)</span>
        <input v-model.number="measurementsCount90" type="number" min="0" step="1" />
      </label>

      <label class="field">
        <span>Количество измерений (50%)</span>
        <input v-model.number="measurementsCount50" type="number" min="0" step="1" />
      </label>

      <label class="field">
        <span>Количество измерений (10%)</span>
        <input v-model.number="measurementsCount10" type="number" min="0" step="1" />
      </label>

      <label class="field">
        <span>Количество измерений (2%)</span>
        <input v-model.number="measurementsCount2" type="number" min="0" step="1" />
      </label>

      <button type="submit" :disabled="!isValid">Выгрузить отчёт в формате xls</button>
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

input[type="text"], input[type="number"], select {
  padding: 10px 12px;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  font-size: 14px;
}

select {
  background-color: white;
  cursor: pointer;
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
