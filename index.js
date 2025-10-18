let loadedCases = []
let loadedGroupValues = {}
const groupDefs = {
	'G1 — основи кібербезпеки': 7,
	'G2 — захист персональних даних': 6,
	'G3 — виявлення кіберзагроз': 5,
	'G4 — підвищення обізнаності та реагування': 4,
}

function computeGroupIntegralScore(scores) {
	return scores.reduce((sum, s) => sum + s, 0)
}

function quadraticSSplineMembership(x, a, b, c) {
	if (a >= b || b >= c) {
		throw new Error('Thresholds must satisfy a < b < c')
	}
	if (x <= a) return 0
	if (x >= c) return 1
	if (x <= b) {
		return Math.pow(x - a, 2) / (2 * Math.pow(b - a, 2))
	}
	return 1 - Math.pow(c - x, 2) / (2 * Math.pow(c - b, 2))
}

function normalizeWeights(weights) {
	const sum = weights.reduce((s, w) => s + w, 0)
	if (sum <= 0) return weights.map(() => 1 / weights.length)
	return weights.map(w => w / sum)
}

function aggregateGroupScore(mu, w) {
	return mu * w
}

function determineLFC(avgTheta) {
	const bounds = [0.8, 0.6, 0.4, 0.2, 0.2]
	const x = Math.max(0, Math.min(1, avgTheta))
	if (x > bounds[0]) return 'Високий рівень'
	if (x > bounds[1]) return 'Вище середнього'
	if (x > bounds[2]) return 'Середній рівень'
	if (x > bounds[3]) return 'Низький рівень'
	return 'Дуже низький рівень'
}

function getThresholds() {
	return [
		[
			parseFloat(document.getElementById('t1a').value),
			parseFloat(document.getElementById('t1b').value),
			parseFloat(document.getElementById('t1c').value),
		],
		[
			parseFloat(document.getElementById('t2a').value),
			parseFloat(document.getElementById('t2b').value),
			parseFloat(document.getElementById('t2c').value),
		],
		[
			parseFloat(document.getElementById('t3a').value),
			parseFloat(document.getElementById('t3b').value),
			parseFloat(document.getElementById('t3c').value),
		],
		[
			parseFloat(document.getElementById('t4a').value),
			parseFloat(document.getElementById('t4b').value),
			parseFloat(document.getElementById('t4c').value),
		],
	]
}

function getWeights() {
	return [
		parseFloat(document.getElementById('w1').value),
		parseFloat(document.getElementById('w2').value),
		parseFloat(document.getElementById('w3').value),
		parseFloat(document.getElementById('w4').value),
	]
}

function compute() {
	if (loadedCases.length === 0) {
		alert('Імпортуйте Excel або згенеруйте демо-таблицю.')
		return
	}

	const groupNames = Object.keys(groupDefs)
	const rawWeights = getWeights()
	const weights = normalizeWeights(rawWeights)
	const thresholds = getThresholds()

	const perCaseDeltas = []
	const perCaseMus = []
	const thetaPerCase = []

	for (const caseName of loadedCases) {
		const deltas = []
		const mus = []

		groupNames.forEach((gname, idx) => {
			const values = loadedGroupValues[caseName][gname]
			const delta = computeGroupIntegralScore(values)
			const [a, b, c] = thresholds[idx]
			const mu = quadraticSSplineMembership(delta, a, b, c)
			deltas.push(delta)
			mus.push(mu)
		})

		perCaseDeltas.push(deltas)
		perCaseMus.push(mus)

		const theta = mus.reduce(
			(sum, mu, i) => sum + aggregateGroupScore(mu, weights[i]),
			0
		)
		thetaPerCase.push(theta)
	}

	const avg = thetaPerCase.reduce((sum, t) => sum + t, 0) / thetaPerCase.length
	const lfc = determineLFC(avg)

	fillCalcTable(perCaseDeltas, perCaseMus, thetaPerCase)

	document.getElementById('avgBadge').textContent = avg.toFixed(3)
	setLFCBadge(lfc)
}

function fillCalcTable(perCaseDeltas, perCaseMus, thetaPerCase) {
	const tbody = document.getElementById('calcTable').querySelector('tbody')
	tbody.innerHTML = ''

	const thead = document.getElementById('calcTable').querySelector('thead tr')
	thead.innerHTML = '<th>Параметр</th>'
	loadedCases.forEach(c => {
		const th = document.createElement('th')
		th.textContent = c
		thead.appendChild(th)
	})

	for (let i = 0; i < 4; i++) {
		let tr = document.createElement('tr')
		tr.innerHTML = `<td>δ${i + 1}</td>`
		perCaseDeltas.forEach(deltas => {
			const td = document.createElement('td')
			td.textContent = deltas[i].toFixed(3)
			tr.appendChild(td)
		})
		tbody.appendChild(tr)

		tr = document.createElement('tr')
		tr.innerHTML = `<td>μ${i + 1}</td>`
		perCaseMus.forEach(mus => {
			const td = document.createElement('td')
			td.textContent = mus[i].toFixed(3)
			tr.appendChild(td)
		})
		tbody.appendChild(tr)
	}

	const tr = document.createElement('tr')
	tr.innerHTML = '<td>ϑ(c)</td>'
	thetaPerCase.forEach(theta => {
		const td = document.createElement('td')
		td.textContent = theta.toFixed(3)
		tr.appendChild(td)
	})
	tbody.appendChild(tr)
}

function setLFCBadge(lfc) {
	const badge = document.getElementById('lfcBadge')
	badge.textContent = lfc
	badge.className = 'summary-value'
}

function importExcel(event) {
	const file = event.target.files[0]
	if (!file) return

	const reader = new FileReader()
	reader.onload = function (e) {
		const data = new Uint8Array(e.target.result)
		const workbook = XLSX.read(data, { type: 'array' })
		const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
		const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })

		processImportedData(jsonData)
	}
	reader.readAsArrayBuffer(file)
}

function processImportedData(data) {
	if (data.length < 2) {
		alert('Файл порожній або неправильного формату')
		return
	}

	const headers = data[0]
	const numericCols = []

	headers.forEach((h, idx) => {
		if (idx > 1 && !isNaN(parseFloat(data[1][idx]))) {
			numericCols.push(idx)
		}
	})

	if (numericCols.length === 0) {
		alert('Не знайдено числових стовпців з оцінками.')
		return
	}

	loadedCases = numericCols.map(idx => headers[idx].toString())
	loadedGroupValues = {}
	loadedCases.forEach(c => {
		loadedGroupValues[c] = {}
		Object.keys(groupDefs).forEach(g => {
			loadedGroupValues[c][g] = []
		})
	})

	const groupNames = Object.keys(groupDefs)
	const groupCounts = Object.values(groupDefs)
	let rowIdx = 1

	groupNames.forEach((gname, gIdx) => {
		const count = groupCounts[gIdx]
		for (let i = 0; i < count && rowIdx < data.length; i++) {
			const row = data[rowIdx]
			numericCols.forEach((colIdx, caseIdx) => {
				const val = parseFloat(row[colIdx]) || 0
				loadedGroupValues[loadedCases[caseIdx]][gname].push(val)
			})
			rowIdx++
		}
	})

	fillInputTable(data, numericCols)
}

function fillInputTable(data, numericCols) {
	const thead = document.getElementById('inputTable').querySelector('thead tr')
	thead.innerHTML = '<th>Група критеріїв</th><th>Критерій</th>'

	numericCols.forEach(idx => {
		const th = document.createElement('th')
		th.textContent = data[0][idx]
		thead.appendChild(th)
	})

	const tbody = document.getElementById('inputTable').querySelector('tbody')
	tbody.innerHTML = ''

	for (let i = 1; i < data.length; i++) {
		const tr = document.createElement('tr')
		const row = data[i]

		const groupCell = document.createElement('td')
		groupCell.textContent = row[0] || ''
		tr.appendChild(groupCell)

		const critCell = document.createElement('td')
		critCell.textContent = row[1] || ''
		tr.appendChild(critCell)

		numericCols.forEach(idx => {
			const td = document.createElement('td')
			td.textContent = parseFloat(row[idx]) || 0
			tr.appendChild(td)
		})

		tbody.appendChild(tr)
	}
}

function generateDemo() {
	const groups = [
		['G1 — основи кібербезпеки', 7],
		['G2 — захист персональних даних', 6],
		['G3 — виявлення кіберзагроз', 5],
		['G4 — підвищення обізнаності та реагування', 4],
	]

	const data = [['Група', 'Критерій', 'c1', 'c2', 'c3', 'c4']]

	groups.forEach(([gname, count]) => {
		for (let i = 0; i < count; i++) {
			data.push([
				gname,
				`K${i + 1}`,
				Math.floor(Math.random() * 10) + 1,
				Math.floor(Math.random() * 10) + 1,
				Math.floor(Math.random() * 10) + 1,
				Math.floor(Math.random() * 10) + 1,
			])
		}
	})

	const ws = XLSX.utils.aoa_to_sheet(data)
	const wb = XLSX.utils.book_new()
	XLSX.utils.book_append_sheet(wb, ws, 'Demo')
	XLSX.writeFile(wb, 'demo_cyber_awareness.xlsx')
}

function exportExcel() {
	const calcTbody = document.getElementById('calcTable').querySelector('tbody')
	if (calcTbody.children.length === 0) {
		alert('Спочатку виконайте обчислення.')
		return
	}

	const data = [['Параметр', ...loadedCases]]
	Array.from(calcTbody.children).forEach(tr => {
		const row = Array.from(tr.children).map(td => td.textContent)
		data.push(row)
	})

	const ws = XLSX.utils.aoa_to_sheet(data)
	const wb = XLSX.utils.book_new()
	XLSX.utils.book_append_sheet(wb, ws, 'Розрахунки')

	const summaryData = [
		['Параметр', 'Значення'],
		['Середній ϑ̅', document.getElementById('avgBadge').textContent],
		['LFC', document.getElementById('lfcBadge').textContent],
	]
	const ws2 = XLSX.utils.aoa_to_sheet(summaryData)
	XLSX.utils.book_append_sheet(wb, ws2, 'Підсумок')

	XLSX.writeFile(wb, 'cyber_awareness_results.xlsx')
}

function openSettings() {
	document.getElementById('settingsModal').style.display = 'block'
}

function closeSettings() {
	document.getElementById('settingsModal').style.display = 'none'
}

function saveSettings() {
	const thresholds = getThresholds()
	for (let i = 0; i < thresholds.length; i++) {
		const [a, b, c] = thresholds[i]
		if (!(a < b && b < c)) {
			alert('Потрібно a < b < c для кожної групи.')
			return
		}
	}
	closeSettings()
	alert('Налаштування збережено')
}

window.onclick = function (event) {
	const modal = document.getElementById('settingsModal')
	if (event.target === modal) {
		closeSettings()
	}
}
