/**
 * Importado ANTES de qualquer outro módulo (primeira linha do main.tsx): instala a leitura
 * descriptografada do localStorage antes que o app ou o Firebase leiam alguma chave.
 */
import { instalarLeituraDescriptografada } from './cryptoStorage';

instalarLeituraDescriptografada();
