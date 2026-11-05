// Gera o PDF do certificado (usado pelo site do participante e pelo admin).
// Precisa do jsPDF carregado antes deste arquivo.

// Opcional: imagem de fundo A4 paisagem (ex.: 'certificado-fundo.png', na mesma pasta). Vazio = moldura padrão.
const CERT_FUNDO = '';

function imagemDataUrl(url) {
  return fetch(url).then(r => { if (!r.ok) throw new Error(); return r.blob(); })
    .then(b => new Promise((ok, erro) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = erro; fr.readAsDataURL(b); }));
}

async function gerarPdf(c) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const W = 297, H = 210;
  let fundo = null;
  if (CERT_FUNDO) { try { fundo = await imagemDataUrl(CERT_FUNDO); } catch (e) {} }
  if (fundo) {
    doc.addImage(fundo, 'PNG', 0, 0, W, H);
  } else {
    doc.setFillColor(11, 61, 145); doc.rect(0, 0, W, 20, 'F');
    doc.setFillColor(245, 183, 0); doc.rect(0, 20, W, 3, 'F');
    doc.setFillColor(11, 61, 145); doc.rect(0, H - 8, W, 8, 'F');
    doc.setDrawColor(11, 61, 145); doc.setLineWidth(0.6); doc.rect(12, 30, W - 24, H - 46);
    doc.setTextColor(255, 255, 255); doc.setFont('helvetica', 'bold'); doc.setFontSize(16);
    doc.text(c.evento.toUpperCase(), W / 2, 13, { align: 'center' });
  }
  doc.setTextColor(11, 61, 145); doc.setFont('helvetica', 'bold'); doc.setFontSize(34);
  doc.text(c.titulo, W / 2, 58, { align: 'center' });
  doc.setTextColor(40, 40, 40); doc.setFont('helvetica', 'normal'); doc.setFontSize(15);
  doc.text(doc.splitTextToSize(c.texto, 225), W / 2, 82, { align: 'center', lineHeightFactor: 1.6 });
  doc.setFontSize(12);
  doc.text(c.localData, W / 2, 145, { align: 'center' });
  doc.setDrawColor(120, 120, 120); doc.setLineWidth(0.3); doc.line(W / 2 - 50, 168, W / 2 + 50, 168);
  doc.setFontSize(11); doc.text(c.assinatura, W / 2, 174, { align: 'center' });
  doc.setFontSize(8); doc.setTextColor(110, 110, 110);
  doc.text('Código de validação: ' + c.codigoValidacao, W / 2, H - 13, { align: 'center' });
  doc.save(c.arquivo);
}
