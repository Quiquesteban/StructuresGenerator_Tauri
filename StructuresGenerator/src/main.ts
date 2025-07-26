import { invoke } from "@tauri-apps/api/core";

let selectedFilePath: string = "";
let sheetNames: string[] = [];

/**
 * Paso 1: Selección de archivo
 */
async function seleccionarArchivo() {
  mostrarPaso(2); // Paso 2: cargando

  const res = await invoke<{
    canceled: boolean;
    file_path?: string;
    sheet_names?: string[];
    error?: string;
  }>("select_excel_file");

  if (res.canceled || res.error || !res.file_path || !res.sheet_names) {
    alert(res.error || "Operación cancelada o inválida.");
    mostrarPaso(1);
    return;
  }

  selectedFilePath = res.file_path;
  sheetNames = res.sheet_names;

  cargarDropdown(sheetNames);
  mostrarPaso(3);
}

/**
 * Paso 3: Selección de hoja
 */
async function procesarHojaSeleccionada() {
  const dropdown = document.querySelector<HTMLSelectElement>("#sheetDropdown");
  if (!dropdown) return;

  const selectedSheet = dropdown.value;
  mostrarPaso(2); // Paso 2: cargando

  try {
    const result = await invoke<any>("process_excel", {
      filePath: selectedFilePath,
      sheetName: selectedSheet,
    });

    console.log("📄 Resultado generado por Python:", result);
    mostrarPaso(4);
    mostrarResultado(result);
  } catch (err) {
    console.error("❌ Error procesando el archivo:", err);
    alert("Error al procesar el Excel");
    mostrarPaso(3);
  }
}

/**
 * Mostrar el resultado (puedes personalizarlo)
 */
function mostrarResultado(result: { file: string; fullPath: string }) {
  const container = document.getElementById("outputCardContainer");
  if (!container) return;

  container.innerHTML = `
  <div
    id="dragCard"
    class="group relative p-4 bg-white rounded-xl shadow-md border border-gray-200 flex items-center gap-3 cursor-pointer hover:shadow-lg transition"
    draggable="true"
  >
    <div class="text-2xl">📄</div>
    <div class="flex-1 overflow-hidden">
      <p class="font-medium text-gray-800 truncate">${result.file}</p>
      <p class="text-xs text-blue-600 hidden group-hover:block">
        Arrastra a TwinCAT o haz clic para abrir
      </p>
    </div>
  </div>
`;

  const card = container.querySelector<HTMLDivElement>("#dragCard");
  if (!card) {
    return;
  }

  // Drag hacia TwinCAT 3 (Windows Explorer)
card.addEventListener("click", async () => {
  try {
    await invoke("abrir_carpeta_con_archivo", {
      path: result.fullPath,
    });
  } catch (err) {
    console.error("❌ Error al abrir carpeta:", err);
    alert("No se pudo abrir la carpeta del archivo.");
  }
});
}

/**
 * Mostrar pasos
 */
function mostrarPaso(num: number) {
  document
    .querySelectorAll(".step")
    .forEach((el) => el.classList.add("hidden"));
  document.querySelector(`#step${num}`)?.classList.remove("hidden");

}

/**
 * Cargar dropdown con las hojas
 */
function cargarDropdown(names: string[]) {
  const dropdown = document.querySelector<HTMLSelectElement>("#sheetDropdown");
  if (!dropdown) return;

  dropdown.innerHTML = "";

  names.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    dropdown.appendChild(option);
  });
}

/**
 * Eventos
 */
window.addEventListener("DOMContentLoaded", () => {
  document
    .getElementById("selectFileBtn")
    ?.addEventListener("click", seleccionarArchivo);

  document
    .getElementById("confirmSheetBtn")
    ?.addEventListener("click", procesarHojaSeleccionada);

  document
    .getElementById("backToStep1From3")
    ?.addEventListener("click", () => mostrarPaso(1));

  document
    .getElementById("backToStep1From4")
    ?.addEventListener("click", () => mostrarPaso(1));

  document
    .getElementById("backToStep3")
    ?.addEventListener("click", () => mostrarPaso(3));
});
