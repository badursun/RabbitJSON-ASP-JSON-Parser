<!--#include file="../../v2/RabbitJSON.v2.asp" -->
<html>
	<head>
		<title>Sample 18 - Large Data Set Performance Test</title>
		<meta charset="utf-8">
		<style>
			body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f7fa; }
			.container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
			.header { text-align: center; margin-bottom: 30px; }
			.stats { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
			.stat-card { flex: 1; min-width: 200px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px; text-align: center; }
			.stat-number { font-size: 2em; font-weight: bold; margin-bottom: 5px; }
			.stat-label { font-size: 0.9em; opacity: 0.9; }
			.table-container { overflow-x: auto; margin-top: 20px; }
			table { width: 100%; border-collapse: collapse; background: white; }
			th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 8px; text-align: left; font-weight: 600; position: sticky; top: 0; }
			td { padding: 10px 8px; border-bottom: 1px solid #eee; }
			tr:hover { background-color: #f8f9ff; }
			.price { font-weight: bold; color: #2d5a87; }
			.stock { padding: 4px 8px; border-radius: 4px; color: white; font-size: 0.85em; font-weight: bold; }
			.stock.high { background: #10b981; }
			.stock.medium { background: #f59e0b; }
			.stock.low { background: #ef4444; }
			.rating { color: #fbbf24; font-weight: bold; }
			.category { background: #e0e7ff; color: #3730a3; padding: 3px 8px; border-radius: 12px; font-size: 0.8em; display: inline-block; }
			.brand { font-weight: 600; color: #374151; }
			.performance { background: #ecfdf5; border: 1px solid #10b981; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
			.performance-title { color: #065f46; font-weight: bold; margin-bottom: 10px; }
		</style>
	</head>
	<body>
		<div class="container">
			<%
			' Performance timing başlat
			Dim startTime, endTime, parseTime, processingTime
			startTime = Timer

			' Large JSON dosyasını yükle
			Set json = CreateRabbitJSON()
			Dim jsonFilePath
			jsonFilePath = Server.MapPath("large-sample-data.json")
			
			' Dosyayı oku ve parse et
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set file = fso.OpenTextFile(jsonFilePath, 1)
			Dim jsonContent
			jsonContent = file.ReadAll()
			file.Close()
			Set file = Nothing
			Set fso = Nothing
			
			' JSON'u parse et
			Dim parseStartTime
			parseStartTime = Timer
			Set result = json.Parse(jsonContent)
			parseTime = Timer - parseStartTime
			
			If result Is Nothing Then
				Response.Write "<div style='color:red; text-align:center;'>"
				Response.Write "<h2>❌ JSON Parse Hatası</h2>"
				Response.Write "<p>" & json.LastError & "</p>"
				Response.Write "</div>"
				Set json = Nothing
				Response.End
			End If
			
			' Processing başlat
			Dim processingStartTime
			processingStartTime = Timer
			
			' Metadata ve summary bilgilerini al
			totalProducts = json.GetValue("metadata.totalProducts", 0)
			lastUpdated = json.GetValue("metadata.lastUpdated", "")
			totalValue = json.GetValue("summary.totalValue", 0)
			averagePrice = json.GetValue("summary.averagePrice", 0)
			
			' Ürün sayısını bul
			Dim productCount, i
			productCount = 0
			Do While json.HasValue("products." & productCount)
				productCount = productCount + 1
			Loop
			
			processingTime = Timer - processingStartTime
			endTime = Timer
			%>
			
			<div class="header">
				<h1>🚀 Large Data Performance Test</h1>
				<p>RabbitJSON v2.1.1 - 50 Ürün Katalogu Test</p>
			</div>
			
			<div class="performance">
				<div class="performance-title">⚡ Performance Metrics</div>
				<strong>JSON Parse Time:</strong> <%= FormatNumber(parseTime * 1000, 2) %> ms<br>
				<strong>Data Processing Time:</strong> <%= FormatNumber(processingTime * 1000, 2) %> ms<br>
				<strong>Total Execution Time:</strong> <%= FormatNumber((endTime - startTime) * 1000, 2) %> ms<br>
				<strong>JSON File Size:</strong> <%= FormatNumber(Len(jsonContent) / 1024, 1) %> KB | <a href="https://rabbitjson.com/v2/all-sample-test/large-sample-data.json" target="_blank">View JSON</a>
			</div>
			
			<div class="stats">
				<div class="stat-card">
					<div class="stat-number"><%= productCount %></div>
					<div class="stat-label">Toplam Ürün</div>
				</div>
				<div class="stat-card">
					<div class="stat-number">₺<%= FormatNumber(totalValue, 0) %></div>
					<div class="stat-label">Toplam Değer</div>
				</div>
				<div class="stat-card">
					<div class="stat-number">₺<%= FormatNumber(averagePrice, 0) %></div>
					<div class="stat-label">Ortalama Fiyat</div>
				</div>
				<div class="stat-card">
					<div class="stat-number"><%= json.Count %></div>
					<div class="stat-label">JSON Keys</div>
				</div>
			</div>
			
			<div class="table-container">
				<table>
					<thead>
						<tr>
							<th>ID</th>
							<th>Ürün Adı</th>
							<th>Kategori</th>
							<th>Marka</th>
							<th>Fiyat</th>
							<th>Stok</th>
							<th>Rating</th>
							<th>Durum</th>
						</tr>
					</thead>
					<tbody>
						<%
						' Tüm ürünleri listele
						For i = 0 To productCount - 1
							productId = json.GetValue("products." & i & ".id", 0)
							productName = json.GetValue("products." & i & ".name", "")
							category = json.GetValue("products." & i & ".category", "")
							brand = json.GetValue("products." & i & ".brand", "")
							price = json.GetValue("products." & i & ".price", 0)
							stock = json.GetValue("products." & i & ".stock", 0)
							rating = json.GetValue("products." & i & ".rating", 0)
							inStock = json.GetValue("products." & i & ".inStock", False)
							
							' Stok durumu renk kodları
							Dim stockClass
							If stock > 30 Then
								stockClass = "high"
							ElseIf stock > 10 Then
								stockClass = "medium"
							Else
								stockClass = "low"
							End If
							
							' Durum metni
							Dim statusText, statusColor
							If inStock Then
								statusText = "✅ Stokta"
								statusColor = "green"
							Else
								statusText = "❌ Tükendi"
								statusColor = "red"
							End If
							
							Response.Write "<tr>"
							Response.Write "<td><strong>" & productId & "</strong></td>"
							Response.Write "<td>" & productName & "</td>"
							Response.Write "<td><span class='category'>" & category & "</span></td>"
							Response.Write "<td><span class='brand'>" & brand & "</span></td>"
							Response.Write "<td class='price'>₺" & FormatNumber(price, 2) & "</td>"
							Response.Write "<td><span class='stock " & stockClass & "'>" & stock & "</span></td>"
							Response.Write "<td class='rating'>⭐ " & rating & "</td>"
							Response.Write "<td style='color:" & statusColor & "'>" & statusText & "</td>"
							Response.Write "</tr>"
						Next
						%>
					</tbody>
				</table>
			</div>
			
			<div style="margin-top: 30px; text-align: center; color: #6b7280; font-size: 0.9em;">
				<p>📊 Test tamamlandı! <%= productCount %> ürün başarıyla parse edildi ve görüntülendi.</p>
				<p>Son güncelleme: <%= lastUpdated %></p>
			</div>
			
			<%
			Set json = Nothing
			%>
		</div>
	</body>
</html> 