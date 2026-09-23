# Patch do App.tsx — Fase 1 (wiring)

Após os arquivos da Fase 1 estarem no repo (branch `refactor/phase1-utils`):

## A. Adicionar imports

```ts
import type { Empresa, EmpresaCredenciada, Premissa, Problema, Diagnostico, Resposta, Solucao, TarefaPlanoAcao, DadosConsultoria, AtividadeCronograma } from './types/domain';
import { cleanDigits, formatCNPJ, isValidCNPJ, formatCEP, formatCPF, isValidCPF } from './lib/formatters/br';
import { getMdaExpirationStatus, parseLocalDate, formatFirestoreDate, getActivityDateStr } from './lib/formatters/dates';
import { sanitizeForFirestore } from './lib/firestore/sanitize';
import { oklchToRgb, captureElementWithHtml2Canvas } from './lib/pdf/capture';
import { CONSULTORIA_AREAS, AREAS, AREAS_ORDER, TIPOS_EMPRESA, IMPACTO_ORDER } from './lib/constants/areas';
import { LICENSE_FIELDS } from './lib/constants/license';
import { normalizeAndFormatArea, globalNormalizedMatch } from './lib/domain/areas';
import { pickLicenseFields, extractItemTimestamp, toJsDate, computeLicenseDaysLeft } from './lib/domain/license';
import { deduplicateRespostas, loadAllLocalRespostas, saveAllLocalRespostas, getRespostasForDiagnostico } from './lib/domain/respostas';
import { isValidLogoSource } from './lib/media/logo';
import { Type } from './lib/ai/schemaTypes';
import { extractAndParseJSON } from './lib/ai/parseJson';
```

## B. Remover blocos duplicados

Apagar definições originais de: Type, CONSULTORIA_AREAS, formatters BR, dates, sanitize, oklch/capture, logo, LICENSE_FIELDS, interfaces de domínio, AREAS/AREAS_ORDER, deduplicateRespostas, extractAndParseJSON, loadAllLocalRespostas.

## C. package.json test script

```json
"test": "tsx test/librarySelection.test.ts && tsx test/licenseDays.test.ts && tsx test/formatters.br.test.ts && tsx test/domain.respostas.test.ts"
```

## D. Validar

```bash
npm test && npm run lint && npm run dev
```
