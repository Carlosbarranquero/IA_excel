# 📊 LlamarOllama: Ejecuta un modelo de IA directamente desde Excel

Este módulo VBA permite enviar un *prompt* desde una celda de Excel a un modelo de lenguaje (LLM) alojado localmente con [Ollama](https://ollama.com), y obtener la respuesta directamente en la hoja de cálculo.

> ✅ No necesitas conexión a internet  
> ✅ No usas ninguna API externa  
> ✅ No necesitas programar

---

## 🚀 ¿Qué hace?

Este módulo define una función personalizada de Excel:  
`=LlamarOllama("Tu pregunta aquí")`  
y devuelve la respuesta del modelo configurado localmente (por defecto `llama3.2`).

---

## 🧠 Requisitos

- [Ollama](https://ollama.com) instalado y ejecutándose en tu ordenador  
- Tener cargado un modelo, por ejemplo:  
  ```bash
  ollama run llama3.2
