/**
 * Confere quais documentos ficaram SEM dono (ownerId vazio, 'local' ou 'local_user').
 * Com as novas regras do Firestore, esses documentos deixam de aparecer para todos.
 *
 * Uso (no seu computador, com a conta de serviço do Firebase):
 *   FIREBASE_SERVICE_ACCOUNT_JSON="$(cat conta-servico.json)" npx tsx scripts/auditar-ownerid.ts
 *   # para atribuir os órfãos a um usuário (ex.: o seu uid de admin):
 *   FIREBASE_SERVICE_ACCOUNT_JSON="$(cat conta-servico.json)" npx tsx scripts/auditar-ownerid.ts --atribuir=SEU_UID
 */
import { getAdminDb } from '../src/server/firebaseAdmin';

const COLLECTIONS = [
  'empresas', 'diagnosticos', 'respostas', 'tarefas_plano', 'empresas_credenciadas',
  'agenda_eventos', 'disc_avaliacoes', 'maturidade_avaliacoes', 'analises_resultado',
  'premissas', 'problemas', 'solucoes', 'areas', 'segmentos',
];
const PLACEHOLDERS = new Set(['', 'local', 'local_user']);

async function main() {
  const db = getAdminDb();
  if (!db) {
    console.error('Defina FIREBASE_SERVICE_ACCOUNT_JSON com o JSON da conta de serviço.');
    process.exit(1);
  }
  const assignArg = process.argv.find((a) => a.startsWith('--atribuir='));
  const assignTo = assignArg ? assignArg.split('=')[1].trim() : '';

  for (const col of COLLECTIONS) {
    const snap = await db.collection(col).get();
    const orphans = snap.docs.filter((d) => {
      const owner = d.get('ownerId');
      return typeof owner !== 'string' || PLACEHOLDERS.has(owner);
    });
    console.log(`${col}: ${snap.size} documento(s), ${orphans.length} sem dono`);
    if (assignTo && orphans.length) {
      for (let i = 0; i < orphans.length; i += 400) {
        const batch = db.batch();
        orphans.slice(i, i + 400).forEach((d) => batch.update(d.ref, { ownerId: assignTo }));
        await batch.commit();
      }
      console.log(`  → ${orphans.length} atribuído(s) a ${assignTo}`);
    }
  }
  process.exit(0);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
