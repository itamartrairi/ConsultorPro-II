import { initializeApp } from 'firebase/app';
import { getFirestore, collection, getDocs, deleteDoc, doc } from 'firebase/firestore';
import { readFileSync } from 'fs';

const config = JSON.parse(readFileSync('./firebase-applet-config.json', 'utf-8'));
const app = initializeApp(config);
const db = getFirestore(app, config.firestoreDatabaseId);

const collectionsToClear = [
  'empresas',
  'diagnosticos',
  'respostas',
  'premissas',
  'problemas',
  'solucoes',
  'areas',
  'tarefas_plano'
];

async function clearData() {
  console.log("Iniciando limpeza de dados experimentais...");
  for (const colName of collectionsToClear) {
    try {
      const snap = await getDocs(collection(db, colName));
      console.log(`Coleção '${colName}': ${snap.size} documento(s) encontrado(s).`);
      for (const d of snap.docs) {
        await deleteDoc(doc(db, colName, d.id));
      }
      console.log(`Coleção '${colName}' limpa com sucesso.`);
    } catch (err) {
      console.error(`Erro ao limpar coleção '${colName}':`, err);
    }
  }
  console.log("Limpeza de dados experimentais concluída com sucesso!");
  process.exit(0);
}

clearData();
