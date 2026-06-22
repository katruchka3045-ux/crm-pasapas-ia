import { createClient } from '@supabase/supabase-js';
import PptxGenJS from 'pptxgenjs';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';

export const config = { maxDuration: 120 };

const SUPABASE_URL = 'https://gjanblhahrauwkfdagbc.supabase.co';
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

const STUDERIA_COLORS = {
  dark:      '1A1040',
  light:     'F5F0FF',
  primary:   '9B59B6',
  secondary: '6C3483',
  accent:    'BB8FCE',
  muted:     'D2B4DE',
  white:     'FFFFFF',
  text:      '2C2040'
};

function slidesCountForDuree(duree) {
  if (duree <= 0.75) return 7;
  if (duree <= 1.0)  return 9;
  if (duree <= 1.5)  return 12;
  return 16;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { sujet, duree, of, nature, chargeType, apiKey } = req.body;

  if (!apiKey || !apiKey.startsWith('sk-')) {
    return res.status(401).json({ error: 'Clé API manquante ou invalide' });
  }
  if (!sujet) {
    return res.status(400).json({ error: 'Sujet manquant' });
  }

  try {
    // 1. Structurer le contenu pédagogique via Claude
    const nbSlides = slidesCountForDuree(duree);
    const dureeMin = Math.round(duree * 60);

    const systemPrompt = `Tu es l'assistante pédagogique de Catherine Le Pennec, formatrice IA pour ${of}.
Tu dois structurer le contenu d'une session de formation de ${dureeMin} minutes sur le thème suivant : "${sujet}".

La session suit TOUJOURS cette structure en 4 parties (sandwich accueil/théorie/pratique/QA) :
1. ACCUEIL & INTRO (court, présentation du thème et du déroulé)
2. THÉORIE (concepts essentiels, chiffres clés si pertinent, structurés en blocs visuels)
3. PRATIQUE (exercices concrets avec l'IA, avec des prompts copier-coller pour les apprenants)
4. QA & CLÔTURE (points clés à retenir + anticipation de questions fréquentes)

Style : ton direct, terrain, pas de jargon abstrait, toujours orienté mise en pratique immédiate.
Réponds UNIQUEMENT en JSON valide, sans markdown, selon ce schéma exact :

{
  "titre": "string court et percutant",
  "sousTitre": "string",
  "dureeAffichee": "${dureeMin} min",
  "programme": [
    {"phase": "Accueil & intro", "minutes": number},
    {"phase": "Théorie", "minutes": number, "detail": "string"},
    {"phase": "Pratique", "minutes": number, "detail": "string"},
    {"phase": "Q&A & clôture", "minutes": number, "detail": "string"}
  ],
  "slides": [
    {
      "partie": "accueil|theorie|pratique|qa",
      "titre": "string",
      "sousTitre": "string optionnel",
      "type": "texte|chiffres|exercice|recap",
      "contenu": ["string", "string"],
      "exercicePrompt": "string optionnel - prompt copier-coller si type=exercice"
    }
  ],
  "pitchOral": [
    {"slideIndex": number, "texte": "ce que Catherine doit dire pour cette slide, ton oral naturel"}
  ],
  "faqAnticipee": [
    {"question": "string", "reponse": "string courte et directe"}
  ]
}

Génère environ ${nbSlides} slides au total répartis sur les 4 parties. Pour faqAnticipee, génère 6 à 10 questions plausibles que les apprenants pourraient poser sur ce sujet précis.`;

    const claudeResponse = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 8000,
        system: systemPrompt,
        messages: [{ role: 'user', content: `Sujet de la session : ${sujet}` }]
      })
    });

    if (!claudeResponse.ok) {
      const errBody = await claudeResponse.json().catch(() => ({}));
      throw new Error('Erreur API Claude : ' + (errBody.error?.message || claudeResponse.status));
    }

    const claudeData = await claudeResponse.json();
    const rawText = claudeData.content?.[0]?.text || '{}';
    const cleanJson = rawText.replace(/```json|```/g, '').trim();
    const contenu = JSON.parse(cleanJson);

    // 2. Générer le PPTX
    const pptxBuffer = await buildPptx(contenu, sujet);

    // 3. Générer le DOCX (pitch oral + FAQ)
    const docxBuffer = await buildPitchDocx(contenu, sujet);

    // 4. Upload vers Supabase Storage
    const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
    const safeSlug = sujet.toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/(^-|-$)/g, '')
      .slice(0, 60);
    const timestamp = Date.now();
    const pptxPath = `${safeSlug}-${timestamp}.pptx`;
    const pitchPath = `${safeSlug}-${timestamp}-pitch.docx`;

    const { error: pptxErr } = await supabase.storage
      .from('supports-cours')
      .upload(pptxPath, pptxBuffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      });
    if (pptxErr) throw new Error('Upload PPTX : ' + pptxErr.message);

    const { error: docxErr } = await supabase.storage
      .from('supports-cours')
      .upload(pitchPath, docxBuffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
    if (docxErr) throw new Error('Upload DOCX : ' + docxErr.message);

    const { data: pptxUrlData } = supabase.storage.from('supports-cours').getPublicUrl(pptxPath);
    const { data: pitchUrlData } = supabase.storage.from('supports-cours').getPublicUrl(pitchPath);

    return res.status(200).json({
      pptxUrl: pptxUrlData.publicUrl,
      pitchUrl: pitchUrlData.publicUrl
    });

  } catch (e) {
    console.error('generate-support error:', e);
    return res.status(500).json({ error: e.message || 'Erreur génération support' });
  }
}

// ════════════════════════════════════════
// CONSTRUCTION DU PPTX
// ════════════════════════════════════════
async function buildPptx(contenu, sujet) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.author = 'Studeria';
  pres.title = contenu.titre || sujet;

  const C = STUDERIA_COLORS;

  // Slide titre
  const s1 = pres.addSlide();
  s1.background = { color: C.dark };
  s1.addText(contenu.titre, {
    x: 0.5, y: 2.3, w: 9, h: 1.2,
    fontSize: 38, fontFace: 'Trebuchet MS', color: C.white, bold: true, align: 'center'
  });
  if (contenu.sousTitre) {
    s1.addText(contenu.sousTitre, {
      x: 0.5, y: 3.4, w: 9, h: 0.6,
      fontSize: 18, fontFace: 'Calibri', color: C.accent, align: 'center'
    });
  }
  s1.addText(contenu.dureeAffichee || '', {
    x: 0.5, y: 4.6, w: 9, h: 0.4,
    fontSize: 14, fontFace: 'Calibri Light', color: C.muted, align: 'center'
  });

  // Slide programme
  const s2 = pres.addSlide();
  s2.background = { color: C.light };
  s2.addText('Au programme', {
    x: 0.5, y: 0.4, w: 9, h: 0.6,
    fontSize: 28, fontFace: 'Trebuchet MS', color: C.primary, bold: true
  });
  let progY = 1.4;
  (contenu.programme || []).forEach(p => {
    s2.addText(p.phase, {
      x: 0.7, y: progY, w: 3.5, h: 0.5,
      fontSize: 16, fontFace: 'Calibri', color: C.text, bold: true
    });
    s2.addText(`${p.minutes} min${p.detail ? ' — ' + p.detail : ''}`, {
      x: 4.3, y: progY, w: 5, h: 0.5,
      fontSize: 13, fontFace: 'Calibri', color: C.secondary
    });
    progY += 0.65;
  });

  // Slides de contenu
  (contenu.slides || []).forEach(sl => {
    const slide = pres.addSlide();
    slide.background = { color: C.light };

    const partieLabel = { accueil: 'ACCUEIL', theorie: 'THÉORIE', pratique: 'PRATIQUE', qa: 'Q&A & CLÔTURE' }[sl.partie] || '';
    slide.addText(partieLabel, {
      x: 0.5, y: 0.3, w: 9, h: 0.35,
      fontSize: 11, fontFace: 'Calibri', color: C.secondary, bold: true
    });
    slide.addText(sl.titre, {
      x: 0.5, y: 0.65, w: 9, h: 0.8,
      fontSize: 26, fontFace: 'Trebuchet MS', color: C.primary, bold: true
    });
    if (sl.sousTitre) {
      slide.addText(sl.sousTitre, {
        x: 0.5, y: 1.35, w: 9, h: 0.4,
        fontSize: 14, fontFace: 'Calibri', color: C.text
      });
    }

    let contentY = sl.sousTitre ? 1.9 : 1.6;
    (sl.contenu || []).forEach(line => {
      slide.addText(line, {
        x: 0.7, y: contentY, w: 8.6, h: 0.5,
        fontSize: 14, fontFace: 'Calibri', color: C.text,
        bullet: true
      });
      contentY += 0.55;
    });

    if (sl.type === 'exercice' && sl.exercicePrompt) {
      slide.addShape(pres.ShapeType.roundRect, {
        x: 0.7, y: contentY + 0.2, w: 8.6, h: 1.8,
        fill: { color: 'FFFFFF' }, line: { color: C.muted, width: 1 }, rectRadius: 0.1
      });
      slide.addText('Prompt à copier-coller :', {
        x: 0.9, y: contentY + 0.35, w: 8.2, h: 0.3,
        fontSize: 11, fontFace: 'Calibri', color: C.secondary, bold: true
      });
      slide.addText(sl.exercicePrompt, {
        x: 0.9, y: contentY + 0.65, w: 8.2, h: 1.2,
        fontSize: 11, fontFace: 'Courier New', color: C.text
      });
    }
  });

  // Slide de fin
  const sEnd = pres.addSlide();
  sEnd.background = { color: C.dark };
  sEnd.addText('Merci pour votre participation !', {
    x: 0.5, y: 2.8, w: 9, h: 0.8,
    fontSize: 30, fontFace: 'Trebuchet MS', color: C.white, bold: true, align: 'center'
  });
  sEnd.addText('Retrouvez le replay sur Softair', {
    x: 0.5, y: 3.7, w: 9, h: 0.5,
    fontSize: 16, fontFace: 'Calibri', color: C.accent, align: 'center'
  });

  return await pres.write('nodebuffer');
}

// ════════════════════════════════════════
// CONSTRUCTION DU DOCX (PITCH ORAL)
// ════════════════════════════════════════
async function buildPitchDocx(contenu, sujet) {
  const children = [];

  children.push(new Paragraph({
    text: `SCRIPT ORAL — ${contenu.titre || sujet}`,
    heading: HeadingLevel.TITLE
  }));
  children.push(new Paragraph({
    text: `Durée : ${contenu.dureeAffichee || ''}`,
    spacing: { after: 300 }
  }));

  (contenu.pitchOral || []).forEach(p => {
    const slideRef = contenu.slides?.[p.slideIndex];
    children.push(new Paragraph({
      text: `SLIDE ${p.slideIndex + 1}${slideRef ? ' — ' + slideRef.titre : ''}`,
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 300 }
    }));
    children.push(new Paragraph({
      children: [new TextRun({ text: p.texte })],
      spacing: { after: 200 }
    }));
  });

  children.push(new Paragraph({
    text: 'QUESTIONS ANTICIPÉES',
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 500 }
  }));

  (contenu.faqAnticipee || []).forEach(qa => {
    children.push(new Paragraph({
      children: [new TextRun({ text: qa.question, bold: true })],
      spacing: { before: 200 }
    }));
    children.push(new Paragraph({
      children: [new TextRun({ text: qa.reponse })],
      spacing: { after: 100 }
    }));
  });

  const doc = new Document({ sections: [{ children }] });
  return await Packer.toBuffer(doc);
}
