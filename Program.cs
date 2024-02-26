using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Newtonsoft.Json;
// using Newtonsoft.Json.Linq;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.WebHost.ConfigureKestrel(options => options.ListenLocalhost(5070));
//builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    //   app.UseSwagger();
    //   app.UseSwaggerUI();
}

app.UseHttpsRedirection();

Dictionary<int, List<string>> tiposModulo = new Dictionary<int, List<string>>
{
    { 1, new List<string> { "json1" } },
    { 2, new List<string> { "json2", "json3", "json4" } },
};

app.MapGet("/api/exportarXlsx", (
    [FromQuery] string etapas
) =>
{
    //transformar array de etapas string em list int
    List<int> etapasInt = etapas.Split(',').Select(int.Parse).ToList();
    //montar o base path para buscar o relatorio
    string path = $"./json_files";

    //percorre as etapas selecionadas
    foreach (var etapa in etapasInt)
    {
        // caso nao exista a etapa no dicionario de modulos continua o loop
        if (!tiposModulo.ContainsKey(etapa))
        {
            continue;
        }

        using (var workbook = new XLWorkbook())
        {
            // percorre os modulos da etapa selecionada
            foreach (var tipo in tiposModulo[etapa])
            {
                // termina de montar o caminho para o arquivo final
                string pathFile = path + $"/{tipo}.json";
                // se o arquivo existir, adiciona na planilha
                if (File.Exists(pathFile))
                {
                    // adiciona uma nova sheet a planilha
                    var worksheet = workbook.Worksheets.Add(tipo);
                    var currentRow = 1;

                    // inicia a leitura do arquivo json
                    List<dynamic> data = new List<dynamic>();
                    // try catch da gambiarra, mas ta valendo, meo!
                    try
                    {
                        // caso o json seja um array, transforma em lista
                        data = JsonConvert.DeserializeObject<List<dynamic>>(File.ReadAllText(pathFile));
                    }
                    catch (System.Exception)
                    {
                        // caso o json seja um objeto, transforma em lista
                        data.Add(JsonConvert.DeserializeObject<dynamic>(File.ReadAllText(pathFile)));
                    }

                    foreach (var objetoJson in data)
                    {
                        // tratamento pra ignorar algum objeto
                        // if (objetoJson["isHeader"] == true || objetoJson["isFooter"] == true || objetoJson["isTotal"] == true)
                        // {
                        //     continue;
                        // }

                        // adiciona o cabecalho na planilha
                        if (currentRow == 1)
                        {
                            int count = 1;
                            // percorre as propriedades do objeto
                            foreach (var item in objetoJson.Properties())
                            {
                                worksheet.Cell(currentRow, count).Value = item.Name.ToString();
                                count++;
                            }
                        }
                        currentRow++;

                        int count2 = 1;
                        // aqui inicia a escrita dos dados
                        foreach (var item in objetoJson)
                        {
                            worksheet.Cell(currentRow, count2).SetValue(item.Value.ToString());
                            count2++;
                        }
                    }
                }
            }

            // Salvar o arquivo fora do loop, após adicionar todas as planilhas
            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                var content = stream.ToArray();

                File.WriteAllBytes("./xlsx_files/relatorio.xlsx", content);
            }
        }

    }


    return etapasInt;
})
.WithOpenApi();

app.Run();

