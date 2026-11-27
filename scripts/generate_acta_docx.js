const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  ImageRun,
  ShadingType,
  PageBreak
} = require('docx');

const baseDir = 'C:/Users/USER/.gemini/antigravity/scratch/municipio-escolar';
const insigniaPath = path.join(baseDir, 'public/assets/insignia.png');
const imgInsignia = fs.readFileSync(insigniaPath);

const borderNone = {
  top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
  right: { style: BorderStyle.NONE, size: 0, color: 'auto' }
};

const borderCell = {
  top: { style: BorderStyle.SINGLE, size: 1, color: '94a3b8' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: '94a3b8' },
  left: { style: BorderStyle.SINGLE, size: 1, color: '94a3b8' },
  right: { style: BorderStyle.SINGLE, size: 1, color: '94a3b8' }
};

const borderThick = {
  top: { style: BorderStyle.SINGLE, size: 2, color: '1e3a8a' },
  bottom: { style: BorderStyle.SINGLE, size: 2, color: '1e3a8a' },
  left: { style: BorderStyle.SINGLE, size: 2, color: '1e3a8a' },
  right: { style: BorderStyle.SINGLE, size: 2, color: '1e3a8a' }
};

async function buildActaDocx() {
  console.log('Generating Official Acta Docx...');

  const docChildren = [
    // Header Table with Insignia
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderNone,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 18, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: imgInsignia,
                      transformation: { width: 80, height: 100 }
                    })
                  ]
                })
              ]
            }),
            new TableCell({
              width: { size: 82, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "REPÚBLICA DEL PERÚ • MINISTERIO DE EDUCACIÓN", bold: true, size: 18, color: "475569", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "DIRECCIÓN REGIONAL DE EDUCACIÓN CAJAMARCA - UGEL CAJAMARCA", bold: true, size: 16, color: "64748b", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "I.E. SAN MARCELINO CHAMPAGNAT", bold: true, size: 28, color: "1e3a8a", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "COMITÉ ELECTORAL ESTUDIANTIL DEL MUNICIPIO ESCOLAR", bold: true, size: 20, color: "b45309", font: "Arial" })
                  ]
                })
              ]
            })
          ]
        })
      ]
    }),

    new Paragraph({
      spacing: { before: 80, after: 120 },
      children: [
        new TextRun({ text: "_________________________________________________________________________________", color: "1e3a8a" })
      ]
    }),

    // Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 80 },
      children: [
        new TextRun({
          text: "ACTA GENERAL DE ESCRUTINIO, CÓMPUTO FINAL Y PROCLAMACIÓN DE RESULTADOS",
          bold: true,
          size: 24,
          color: "0f172a",
          font: "Arial"
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [
        new TextRun({
          text: "ELECCIONES DEMOCRÁTICAS DEL MUNICIPIO ESCOLAR 2026 - 2027",
          bold: true,
          size: 20,
          color: "1e3a8a",
          font: "Arial"
        })
      ]
    }),

    // Introduction paragraph
    new Paragraph({
      alignment: AlignmentType.JUSTIFY,
      spacing: { after: 140 },
      children: [
        new TextRun({
          text: "En la ciudad de Cajamarca, a los 24 días del mes de noviembre del año 2026, siendo las 17:00 horas, en los ambientes de la Institución Educativa San Marcelino Champagnat (Turno Tarde), se reunieron los miembros del Comité Electoral Estudiantil conjuntamente con los docentes asesores electorales y los personeros acreditados de las listas participantes, con el objeto de formalizar el escrutinio general, cómputo definitivo de votos y la proclamación de la lista ganadora que conformará el Municipio Escolar para el periodo lectivo correspondiente.",
          size: 19,
          font: "Arial"
        })
      ]
    }),

    // I. DATOS GENERALES DEL PROCESO
    new Paragraph({
      spacing: { before: 100, after: 80 },
      children: [
        new TextRun({ text: "I. DATOS GENERALES DEL PROCESO ELECTORAL", bold: true, size: 20, color: "1e3a8a", font: "Arial" })
      ]
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderCell,
      rows: [
        new TableRow({
          children: [
            new TableCell({ width: { size: 60, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "f1f5f9" }, children: [new Paragraph({ children: [new TextRun({ text: "CONCEPTO", bold: true, size: 17, font: "Arial" })] })] }),
            new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "f1f5f9" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CANTIDAD / VALOR", bold: true, size: 17, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total de Estudiantes Habilitados (Padrón Oficial SIAGIE)", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "589 alumnos", bold: true, size: 17, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total de Estudiantes que Sufragaron (Votos Registrados)", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "524 alumnos (89.0%)", bold: true, size: 17, color: "1e3a8a", font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Total de Estudiantes que No Sufragaron (Ausentismo)", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "65 alumnos (11.0%)", size: 17, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Modalidad de Sufragio", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Voto Electrónico Presencial", size: 17, font: "Arial" })] })] })
          ]
        })
      ]
    }),

    // II. CUADRO DE ESCRUTINIO Y RESULTADOS
    new Paragraph({
      spacing: { before: 180, after: 80 },
      children: [
        new TextRun({ text: "II. CUADRO OFICIAL DE RESULTADOS DE VOTACIÓN", bold: true, size: 20, color: "1e3a8a", font: "Arial" })
      ]
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderCell,
      rows: [
        new TableRow({
          children: [
            new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "N°", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ width: { size: 36, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "LISTA CANDIDATA", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ width: { size: 36, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "CANDIDATO(A) A ALCALDE(SA)", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VOTOS", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ width: { size: 10, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "%", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] })
          ]
        }),
        // Lista 2 - Ganadora
        new TableRow({
          children: [
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "ecfdf5" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "2", bold: true, color: "047857", size: 18, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "ecfdf5" }, children: [new Paragraph({ children: [new TextRun({ text: "Movimiento Revolucionario Champagnino", bold: true, color: "047857", size: 17, font: "Arial" }), new TextRun({ text: " (Renovación champagnina)", size: 14, color: "065f46", font: "Arial" }), new TextRun({ text: "  [GANADORA]", bold: true, size: 14, color: "059669", font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "ecfdf5" }, children: [new Paragraph({ children: [new TextRun({ text: "Jhostin Aurelio Anchay Vega\n", bold: true, size: 16, font: "Arial" }), new TextRun({ text: "Teniente: Sergio Manuel Soto Arbildo (4to \"C\")", size: 14, color: "475569", font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "ecfdf5" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "281", bold: true, color: "047857", size: 18, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "ecfdf5" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "55.0 %", bold: true, color: "047857", size: 17, font: "Arial" })] })] })
          ]
        }),
        // Lista 3
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3", bold: true, color: "ea580c", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Unión y Liderazgo", bold: true, color: "ea580c", size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Thiago Jhean pool campos Fernández\n", bold: true, size: 16, font: "Arial" }), new TextRun({ text: "Teniente: Marco Antonio Huaccha Cotrina (2do B)", size: 14, color: "475569", font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "117", bold: true, size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "22.9 %", size: 16, font: "Arial" })] })] })
          ]
        }),
        // Lista 1
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1", bold: true, color: "2563eb", size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Legado Marcelino", bold: true, color: "2563eb", size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Luis Augusto Muñoz Chuquilín\n", bold: true, size: 16, font: "Arial" }), new TextRun({ text: "Teniente: Paulo André Ventura Incio (4to B)", size: 14, color: "475569", font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "113", bold: true, size: 17, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "22.1 %", size: 16, font: "Arial" })] })] })
          ]
        }),
        // Totales Válidos
        new TableRow({
          children: [
            new TableCell({ colSpan: 3, shading: { type: ShadingType.CLEAR, fill: "f8fafc" }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "TOTAL DE VOTOS VÁLIDOS:", bold: true, size: 17, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "f8fafc" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "511", bold: true, size: 18, color: "1e3a8a", font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "f8fafc" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "100.0 %", bold: true, size: 16, font: "Arial" })] })] })
          ]
        }),
        // Votos en Blanco (Opcional según registro)
        new TableRow({
          children: [
            new TableCell({ colSpan: 3, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Votos en Blanco:", size: 15, font: "Arial", color: "64748b" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "0 *", size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "0.0 %", size: 15, font: "Arial" })] })] })
          ]
        }),
        // Votos Nulos (Opcional según registro)
        new TableRow({
          children: [
            new TableCell({ colSpan: 3, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Votos Nulos:", size: 15, font: "Arial", color: "64748b" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "0 *", size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "0.0 %", size: 15, font: "Arial" })] })] })
          ]
        }),
        // Gran Total
        new TableRow({
          children: [
            new TableCell({ colSpan: 3, shading: { type: ShadingType.CLEAR, fill: "e2e8f0" }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "TOTAL GENERAL DE SUFRAGIOS COMPUTADOS:", bold: true, size: 17, color: "0f172a", font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "e2e8f0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "511", bold: true, size: 19, color: "1e3a8a", font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "e2e8f0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "100.0 %", bold: true, size: 17, font: "Arial" })] })] })
          ]
        })
      ]
    }),

    new Paragraph({
      spacing: { before: 80, after: 120 },
      children: [
        new TextRun({
          text: "* Nota del Comité: De acuerdo con la disposición de la mesa electoral y la confirmación final de escrutinio, el cómputo oficial se fundamenta en la totalidad de los 511 votos válidos emitidos hacia las listas contendientes.",
          size: 14,
          color: "64748b",
          italics: true,
          font: "Arial"
        })
      ]
    }),

    // III. PROCLAMACIÓN OFICIAL
    new Paragraph({
      spacing: { before: 140, after: 80 },
      children: [
        new TextRun({ text: "III. PROCLAMACIÓN DE LA LISTA GANADORA", bold: true, size: 20, color: "1e3a8a", font: "Arial" })
      ]
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderThick,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, fill: "f0fdf4" },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 120, after: 60 },
                  children: [
                    new TextRun({ text: "¡LISTA GANADORA POR MAYORÍA ABSOLUTA!", bold: true, size: 22, color: "047857", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 40 },
                  children: [
                    new TextRun({ text: "LISTA N° 2: \"MOVIMIENTO REVOLUCIONARIO CHAMPAGNINO\"", bold: true, size: 24, color: "065f46", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 80 },
                  children: [
                    new TextRun({ text: "(Renovación champagnina) — Lema: \"Desarrollo, Unión, Respeto.\"", italics: true, size: 17, color: "047857", font: "Arial" })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 120 },
                  children: [
                    new TextRun({ text: "Con un contundente respaldo de ", size: 18, font: "Arial" }),
                    new TextRun({ text: "281 votos (55.0% de los votos válidos)", bold: true, size: 19, color: "047857", font: "Arial" }),
                    new TextRun({ text: ", habiendo alcanzado la máxima votación requerida por el reglamento electoral escolar.", size: 18, font: "Arial" })
                  ]
                })
              ]
            })
          ]
        })
      ]
    }),

    // IV. CUADRO DE AUTORIDADES ESCOLARES ELECTAS
    new Paragraph({
      spacing: { before: 160, after: 80 },
      children: [
        new TextRun({ text: "IV. INTEGRANTES ELECTOS DEL MUNICIPIO ESCOLAR 2026 - 2027", bold: true, size: 20, color: "1e3a8a", font: "Arial" })
      ]
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderCell,
      rows: [
        new TableRow({
          children: [
            new TableCell({ width: { size: 28, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "047857" }, children: [new Paragraph({ children: [new TextRun({ text: "CARGO ELECTO", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
            new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "047857" }, children: [new Paragraph({ children: [new TextRun({ text: "APELLIDOS Y NOMBRES", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
            new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "047857" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DNI", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
            new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: { type: ShadingType.CLEAR, fill: "047857" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "GRADO / SECC.", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Alcalde Escolar", bold: true, size: 15, color: "047857", font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jhostin Aurelio Anchay Vega", bold: true, size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "62485011", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "4to \"C\"", bold: true, size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Teniente Alcalde", bold: true, size: 15, color: "047857", font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Sergio Manuel Soto Arbildo", bold: true, size: 16, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "75640775", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "4to \"C\"", bold: true, size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Regidor de Educación", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jose David Aceijas Galvez", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "77300942", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3ro \"E\"", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Regidor de Derechos", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Manuel Del Campo Medina", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "77000271", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3ro \"E\"", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Regidor de Bienestar", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Sebastian Zuñiga Ambrosio", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "62535012", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3ro \"E\"", size: 15, font: "Arial" })] })] })
          ]
        })
      ]
    }),

    // V. CIERRE Y FIRMAS
    new Paragraph({
      spacing: { before: 180, after: 140 },
      alignment: AlignmentType.JUSTIFY,
      children: [
        new TextRun({
          text: "No habiendo más asuntos que tratar, y en señal de absoluta conformidad, transparencia y legalidad del proceso electoral, firman la presente Acta de Escrutinio y Proclamación los miembros del Comité Electoral, docentes veedores y los personeros acreditados de las listas participantes, dándose por concluido el acto electoral.",
          size: 18,
          font: "Arial"
        })
      ]
    }),

    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderNone,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Presidente del Comité Electoral\nDNI: ____________________", size: 15, font: "Arial" })] })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Secretario del Comité Electoral\nDNI: ____________________", size: 15, font: "Arial" })] })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Vocal del Comité Electoral\nDNI: ____________________", size: 15, font: "Arial" })] })
              ]
            })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ spacing: { before: 180 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Docente Asesor Electoral\nDNI: ____________________", size: 15, font: "Arial" })] })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ spacing: { before: 180 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Personero Lista 2 (Ganadora)\nDNI: ____________________", size: 15, font: "Arial" })] })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({ spacing: { before: 180 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________", color: "64748b" })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Dirección / Subdirección\nI.E. San Marcelino Champagnat", size: 15, font: "Arial" })] })
              ]
            })
          ]
        })
      ]
    })
  ];

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 720,
              right: 720,
              bottom: 720,
              left: 720
            }
          }
        },
        children: docChildren
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);

  // Save to desktop
  const desktopPath = 'C:/Users/USER/Desktop/ACTA_DE_ESCRUTINIO_OFICIAL_2026.docx';
  fs.writeFileSync(desktopPath, buffer);
  console.log('Saved to Desktop:', desktopPath);

  // Save to Artifacts
  const artifactPath = 'C:/Users/USER/.gemini/antigravity/brain/3f337958-0896-4525-b754-a00032c8cd98/ACTA_DE_ESCRUTINIO_OFICIAL_2026.docx';
  fs.writeFileSync(artifactPath, buffer);
  console.log('Saved to Artifacts:', artifactPath);
}

buildActaDocx().then(() => {
  console.log('Acta DOCX generated successfully!');
}).catch(err => {
  console.error('Error generating Acta DOCX:', err);
  process.exit(1);
});
