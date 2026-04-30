# ILLANES

Proyecto Classic ASP para gestión logística y de reclamos.

## ChatGPT Cloud
Esta sección explica cómo usar ChatGPT Cloud con este proyecto en tres niveles:

1. **Uso web**
2. **Integración API**
3. **Copilot Chat en VS Code**

---

## 1. Usar ChatGPT en la web
1. Abre `https://chat.openai.com`
2. Inicia sesión con tu cuenta OpenAI / GPT
3. Si necesitas funciones avanzadas, suscríbete a ChatGPT Plus / Enterprise

> Usa ChatGPT para revisar código, generar documentación y planificar cambios.

---

## 2. Integración de la API de OpenAI
### Pasos
1. Regístrate en `https://platform.openai.com`
2. Genera una clave de API
3. Crea un archivo `.env` en la raíz del proyecto con:

```text
OPENAI_API_KEY=tu_api_key_aqui
```

4. Usa el ejemplo en `openai_chatgpt_example.asp` para hacer llamadas desde Classic ASP.

### Ejemplo de archivo de entorno
Copia `.env.example` a `.env` y reemplaza la clave.

---

## 3. Copilot Chat en VS Code
### Instalación
1. Abre VS Code
2. Instala la extensión `GitHub Copilot` o `GitHub Copilot Chat`
3. Inicia sesión con tu cuenta GitHub / OpenAI

### Uso
- Usa sugerencias mientras editas archivos ASP.
- Abre el panel de Copilot Chat para hacer preguntas sobre el código.

---

## Archivos añadidos
- `.env.example`: plantilla para la clave de OpenAI
- `openai_chatgpt_example.asp`: ejemplo de llamada a la API de ChatGPT desde Classic ASP

---

## Notas de seguridad
- No subas tu `.env` al repositorio.
- El archivo `.gitignore` ya excluye `.env` y `*.env`.
