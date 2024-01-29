const { google } = require('googleapis');
const credentials = require('./credentials.json');

const docId = '1WDDbNHZEq29PFYjtF6_A36ZxdueXZnByxyn2KU8C8NE';

// Configurando a autenticação usando as credenciais
const auth = new google.auth.GoogleAuth({
  credentials: credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
});

// Criando uma instância do Google Sheets API
const sheets = google.sheets({ version: 'v4', auth });

// Obtendo informações da planilha
sheets.spreadsheets.get({
  spreadsheetId: docId,
}).then(response => {
  const sheetTitle = response.data.properties.title;
  const studentsSheet = response.data.sheets[0];

  const range = `${studentsSheet.properties.title}!A2:H`;

  sheets.spreadsheets.values.get({
    spreadsheetId: docId,
    range: range,
  }).then(studentData => {
    const rows = studentData.data.values;

    if (rows.length) {
      // Colunas: Matricula, Aluno, Faltas, P1, P2, P3, Situação, Nota para Aprovação Final
      rows.forEach(row => {
        const matricula = row[0];
        const aluno = row[1];
        const faltas = parseInt(row[2]);
        const p1 = parseFloat(row[3]);
        const p2 = parseFloat(row[4]);
        const p3 = parseFloat(row[5]);

        if (isNaN(faltas) || isNaN(p1) || isNaN(p2) || isNaN(p3)) {
          console.error(`Dados inválidos para o aluno: ${aluno}`);
          return;
        }

        const media = (p1 + p2 + p3) / 3;

        console.log(`Calculando situação para o aluno: ${aluno}, Média: ${media}`);

        if (faltas > 15) {
          console.log(`Situação para o aluno ${aluno}: Reprovado por Falta`);
        } else if (media < 5) {
          console.log(`Situação para o aluno ${aluno}: Reprovado por Nota`);
        } else if (media >= 5 && media < 7) {
          const naf = Math.ceil((2 * 5) - media); // Nota para Aprovação Final (naf)
          console.log(`Situação para o aluno ${aluno}: Exame Final, Nota para Aprovação Final (naf): ${naf}`);
        } else {
          console.log(`Situação para o aluno ${aluno}: Aprovado`);
        }
      });
    } else {
      console.log('Nenhum dado de aluno encontrado.');
    }
  }).catch(err => {
    console.error('Erro ao obter dados dos alunos:', err.message);
  });

}).catch(err => {
  console.error('Erro ao obter informações da planilha:', err.message);
});
