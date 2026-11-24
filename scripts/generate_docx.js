const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
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
const candidates = JSON.parse(fs.readFileSync(path.join(baseDir, 'data/candidates.json'), 'utf8'));

const assetsDir = path.join(baseDir, 'public/assets');
const outputDir = path.join(assetsDir, 'generated');
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

const insigniaPath = path.join(assetsDir, 'insignia.png');

// Use the uploaded card images
const cardImg1 = fs.readFileSync(path.join(assetsDir, 'candidates/card_lista_1.jpg'));
const cardImg2 = fs.readFileSync(path.join(assetsDir, 'candidates/card_lista_2.jpg'));
const cardImg3 = fs.readFileSync(path.join(assetsDir, 'candidates/card_lista_3.jpg'));
const cardImages = [cardImg1, cardImg2, cardImg3];

async function buildDocx() {
  console.log('Building Word Document (.docx)...');

  const imgInsignia = fs.readFileSync(insigniaPath);

  const borderNone = {
    top: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    bottom: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    left: { style: BorderStyle.NONE, size: 0, color: 'auto' },
    right: { style: BorderStyle.NONE, size: 0, color: 'auto' }
  };

  const borderCell = {
    top: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
    left: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
    right: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' }
  };

  function createHeader() {
    return [
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
                        transformation: { width: 75, height: 95 }
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
                      new TextRun({ text: "I.E. SAN MARCELINO CHAMPAGNAT", bold: true, size: 28, color: "1e3a8a", font: "Arial" })
                    ]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ text: "Cajamarca • Turno Tarde", bold: true, size: 20, color: "475569", font: "Arial" })
                    ]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ text: "COMITÉ ELECTORAL DEL MUNICIPIO ESCOLAR 2026 - 2027", bold: true, size: 22, color: "b45309", font: "Arial" })
                    ]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({ text: "CÉDULA OFICIAL Y FICHA DE LISTAS CANDIDATAS", bold: true, size: 18, color: "0f172a", font: "Arial" })
                    ]
                  })
                ]
              })
            ]
          })
        ]
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "_________________________________________________________________________________", color: "94a3b8" })
        ]
      })
    ];
  }

  const docChildren = [];

  // =============================================================
  // PAGE 1: RESUMEN GENERAL DE CÉDULA Y LAS 3 LISTAS
  // =============================================================
  docChildren.push(...createHeader());

  docChildren.push(
    new Paragraph({ spacing: { before: 140, after: 100 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "CÉDULA DE VOTACIÓN OFICIAL Y CUADRO DE CANDIDATOS",
          bold: true,
          size: 22,
          color: "1e293b",
          font: "Arial"
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "Listas inscritas para las Elecciones del Municipio Escolar 2026 - 2027:",
          size: 18,
          color: "475569",
          italics: true,
          font: "Arial"
        })
      ]
    }),
    new Paragraph({ spacing: { before: 100, after: 100 } })
  );

  // Table with 3 lists images side-by-side
  docChildren.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderNone,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: cardImages[0],
                      transformation: { width: 180, height: 240 }
                    })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "LISTA 1\n", bold: true, size: 16, color: "2563eb", font: "Arial" }),
                    new TextRun({ text: "Legado Marcelino", bold: true, size: 15, font: "Arial" })
                  ]
                })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: cardImages[1],
                      transformation: { width: 180, height: 240 }
                    })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "LISTA 2\n", bold: true, size: 16, color: "059669", font: "Arial" }),
                    new TextRun({ text: "Renovación champagnina", bold: true, size: 15, font: "Arial" })
                  ]
                })
              ]
            }),
            new TableCell({
              width: { size: 33, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: cardImages[2],
                      transformation: { width: 180, height: 240 }
                    })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({ text: "LISTA 3\n", bold: true, size: 16, color: "ea580c", font: "Arial" }),
                    new TextRun({ text: "Unión y Liderazgo", bold: true, size: 15, font: "Arial" })
                  ]
                })
              ]
            })
          ]
        })
      ]
    }),
    new Paragraph({ spacing: { before: 140, after: 60 } }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: borderCell,
      rows: [
        new TableRow({
          children: [
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "OPCIÓN / LISTA", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "CANDIDATO(A) A ALCALDE(SA)", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "TENIENTE ALCALDE(SA)", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "1e3a8a" }, children: [new Paragraph({ children: [new TextRun({ text: "GRADO / SECC.", bold: true, color: "ffffff", size: 16, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "LISTA 1: Legado Marcelino", bold: true, color: "2563eb", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Luis Augusto Muñoz Chuquilín", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Paulo André Ventura Incio", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "4to de Secundaria - B", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "LISTA 2: Renovación champagnina", bold: true, color: "059669", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Jhostin Anchav Vega", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Sergio Soto Arbildo", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "4to de secundaria - C", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "LISTA 3: Unión y Liderazgo", bold: true, color: "ea580c", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Thiago Jhean pool campos Fernández", bold: true, size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Marco Antonio Huaccha Cotrina", size: 15, font: "Arial" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "2DO \"B\"", size: 15, font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "f1f5f9" }, children: [new Paragraph({ children: [new TextRun({ text: "VOTO EN BLANCO", bold: true, color: "475569", size: 15, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "f1f5f9" }, colSpan: 3, children: [new Paragraph({ children: [new TextRun({ text: "Opción democrática sin preferencia por ninguna lista electoral.", size: 14, color: "64748b", font: "Arial" })] })] })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "fef2f2" }, children: [new Paragraph({ children: [new TextRun({ text: "VOTO NULO", bold: true, color: "dc2626", size: 15, font: "Arial" })] })] }),
            new TableCell({ shading: { type: ShadingType.CLEAR, fill: "fef2f2" }, colSpan: 3, children: [new Paragraph({ children: [new TextRun({ text: "Voto viciado intencionalmente por el elector.", size: 14, color: "64748b", font: "Arial" })] })] })
          ]
        })
      ]
    }),
    new Paragraph({ spacing: { before: 120, after: 60 } }),
    new Paragraph({ children: [new PageBreak()] })
  );

  // =============================================================
  // PAGES 2, 3, 4: FICHA DE CADA LISTA CON SUS FOTOS REALES
  // =============================================================
  candidates.forEach((cand, idx) => {
    docChildren.push(...createHeader());

    const hexColor = cand.color.replace('#', '');

    docChildren.push(
      new Paragraph({ spacing: { before: 120, after: 80 } }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `HOJA DE REGISTRO: LISTA N° ${cand.number} - "${cand.name.toUpperCase()}"`,
            bold: true,
            size: 24,
            color: hexColor,
            font: "Arial"
          })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `CÉDULA ELECTORAL 2026 - 2027 • COLOR ASIGNADO: ${cand.number === 1 ? 'AZUL' : cand.number === 2 ? 'VERDE' : 'NARANJA / ÁMBAR'}`,
            bold: true,
            size: 17,
            color: "64748b",
            font: "Arial"
          })
        ]
      }),
      new Paragraph({ spacing: { before: 80, after: 80 } })
    );

    // Left = Image, Right = Candidate info box
    docChildren.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: borderNone,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new ImageRun({
                        data: cardImages[idx],
                        transformation: { width: 250, height: 330 }
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                borders: borderCell,
                shading: { type: ShadingType.CLEAR, fill: "f8fafc" },
                children: [
                  new Paragraph({
                    spacing: { before: 60, after: 40 },
                    children: [
                      new TextRun({ text: "DETALLES OFICIALES DE LA LISTA:", bold: true, size: 16, color: hexColor, font: "Arial" })
                    ]
                  }),
                  new Paragraph({
                    children: [
                      new TextRun({ text: "• Nombre de Lista: ", bold: true, size: 15, font: "Arial" }),
                      new TextRun({ text: cand.name, size: 15, font: "Arial" })
                    ]
                  }),
                  cand.slogan ? new Paragraph({
                    children: [
                      new TextRun({ text: "• Lema de Campaña:\n", bold: true, size: 15, font: "Arial" }),
                      new TextRun({ text: `"${cand.slogan}"`, italics: true, size: 15, color: hexColor, font: "Arial" })
                    ]
                  }) : new Paragraph({}),
                  new Paragraph({
                    spacing: { before: 80 },
                    children: [
                      new TextRun({ text: "CANDIDATO(A) A ALCALDE(SA):\n", bold: true, size: 15, color: "1e3a8a", font: "Arial" }),
                      new TextRun({ text: cand.alcalde.name, bold: true, size: 16, font: "Arial" }),
                      new TextRun({ text: ` (${cand.alcalde.grade})\n`, size: 14, color: "64748b", font: "Arial" }),
                      new TextRun({ text: "DNI: _______________________", size: 15, font: "Arial" })
                    ]
                  }),
                  new Paragraph({
                    spacing: { before: 80 },
                    children: [
                      new TextRun({ text: "TENIENTE ALCALDE(SA):\n", bold: true, size: 15, color: "1e3a8a", font: "Arial" }),
                      new TextRun({ text: cand.teniente.name, bold: true, size: 16, font: "Arial" }),
                      new TextRun({ text: ` (${cand.teniente.grade})\n`, size: 14, color: "64748b", font: "Arial" }),
                      new TextRun({ text: "DNI: _______________________", size: 15, font: "Arial" })
                    ]
                  })
                ]
              })
            ]
          })
        ]
      }),
      new Paragraph({ spacing: { before: 140, after: 60 } }),
      new Paragraph({
        children: [
          new TextRun({ text: "CUADRO DE INTEGRANTES Y CONFORMIDAD:", bold: true, size: 18, color: "1e3a8a", font: "Arial" })
        ]
      }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: borderCell,
        rows: [
          new TableRow({
            children: [
              new TableCell({ shading: { type: ShadingType.CLEAR, fill: hexColor }, children: [new Paragraph({ children: [new TextRun({ text: "CARGO", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
              new TableCell({ shading: { type: ShadingType.CLEAR, fill: hexColor }, children: [new Paragraph({ children: [new TextRun({ text: "APELLIDOS Y NOMBRES", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
              new TableCell({ shading: { type: ShadingType.CLEAR, fill: hexColor }, children: [new Paragraph({ children: [new TextRun({ text: "GRADO Y SECC.", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
              new TableCell({ shading: { type: ShadingType.CLEAR, fill: hexColor }, children: [new Paragraph({ children: [new TextRun({ text: "DNI", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] }),
              new TableCell({ shading: { type: ShadingType.CLEAR, fill: hexColor }, children: [new Paragraph({ children: [new TextRun({ text: "FIRMA", bold: true, color: "ffffff", size: 15, font: "Arial" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Alcalde(sa)", bold: true, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cand.alcalde.name, bold: true, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cand.alcalde.grade, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "____________", size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "____________", size: 15, font: "Arial" })] })] })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Teniente Alcalde(sa)", bold: true, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cand.teniente.name, bold: true, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cand.teniente.grade, size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "____________", size: 15, font: "Arial" })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "____________", size: 15, font: "Arial" })] })] })
            ]
          })
        ]
      }),
      new Paragraph({ spacing: { before: 180, after: 40 } }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: borderNone,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________________", color: "64748b" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Firma del Candidato(a) a Alcalde(sa)\nDNI: _______________________", size: 15, font: "Arial" })] })
                ]
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "___________________________________", color: "64748b" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Firma del Personero / Docente Asesor\nDNI: _______________________", size: 15, font: "Arial" })] })
                ]
              })
            ]
          })
        ]
      })
    );

    if (idx < candidates.length - 1) {
      docChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });

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
  const desktopDocx = 'C:/Users/USER/Desktop/FICHA_DE_DATOS_TRES_LISTAS_2026.docx';
  fs.writeFileSync(desktopDocx, buffer);
  console.log('Saved to Desktop:', desktopDocx);

  // Save to project docs folder
  const projectDocx = path.join(baseDir, 'data/FICHA_DE_DATOS_TRES_LISTAS_2026.docx');
  fs.writeFileSync(projectDocx, buffer);
  console.log('Saved to project:', projectDocx);

  // Save to artifact folder
  const artifactDocx = 'C:/Users/USER/.gemini/antigravity/brain/3f337958-0896-4525-b754-a00032c8cd98/FICHA_DE_DATOS_TRES_LISTAS_2026.docx';
  fs.writeFileSync(artifactDocx, buffer);
  console.log('Saved to artifacts:', artifactDocx);
}

buildDocx().then(() => {
  console.log('Word document updated successfully!');
}).catch(err => {
  console.error('Error generating docx:', err);
  process.exit(1);
});
