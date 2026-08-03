const fs = require('fs');
const data = `030. Ciputat	Q1	06/01/2026	Done	75	100
033. Taman Mini	Q1	07/01/2026	Done	88	100
036. Jagakarsa	Q1	07/01/2026	Done	77	100
027. Rawamangun	Q1	08/01/2026	Done	78	100
032. Cawang	Q1	08/01/2026	Done	84	100
024. Grand Wisata	Q1	13/01/2026	Done	83	99
051. Bintaro Sektor 9	Q1	13/01/2026	Done	81	100
026. Tangerang City	Q1	14/01/2026	Done	88	100
044. Depok GDC	Q1	14/01/2026	Done	69	100
008. PIK	Q1	15/01/2026	Done	88	100
053. Alam Sutera	Q1	15/01/2026	Done	82	100
042. Muara Karang	Q1	20/01/2026	Done	84	100
028. Taman Palem	Q1	21/01/2026	Done	90	77
011. Kuningan	Q1	27/01/2026	Done	84	100
018. Harapan Indah	Q1	28/01/2026	Done	72	100
041. Karawaci	Q1	28/01/2026	Done	77	100
038. Greenlake	Q1	29/01/2026	Done	87	100
034. Puri Indah	Q1	03/02/2026	Done	75	100
055. Gajah Mada	Q1	03/02/2026	Done	84	100
040. Bekasi Timur	Q1	04/02/2026	Done	94	100
047. Bogor Tajur	Q1	04/02/2026	Done	94	100
014. Sunter	Q1	05/02/2026	Done	77	92
029. Karawang	Q1	05/02/2026	Done	69	100
031. Sawangan	Q1	10/02/2026	Done	85	100
048. Pik 2	Q1	10/02/2026	Done	81	100
006. Kelapa Gading MOI	Q1	11/02/2026	Done	82	100
022. Cikarang	Q1	11/02/2026	Done	82	100
015. Gading Serpong	Q1	12/02/2026	Done	76	100
035. Ciledug	Q1	12/02/2026	Done	67	75
025. Bogor Baru	Q1	18/02/2026	Done	90	100
019. Cinere	Q1	19/02/2026	Done	85	100
054. Bogor Gunung Batu	Q1	19/02/2026	Done	85	100
045. FX Mall	Q1	24/02/2026	Done	83	100
056. Citra Garden 2	Q1	24/02/2026	Done	67	100
021. Kelapa Gading Boulevard	Q1	25/02/2026	Done	78	100
050. Cibinong	Q1	25/02/2026	Done	77	100
043. Tebet	Q1	26/02/2026	Done	68	100
057. Cikupa	Q1	26/02/2026	Done	88	100
001. Pondok Bambu	Q1	03/03/2026	Done	66	100
052. BSD Rawa Buntu	Q1	03/03/2026	Done	77	100
002. Kemang	Q1	04/03/2026	Done	73	80
007. Bekasi	Q1	04/03/2026	Done	64	100
005. Greenville	Q1	05/03/2026	Done	84	100
010. Bogor	Q1	05/03/2026	Done	87	100
037. Bali	Q1	06/03/2026	Done	92	100
004. Cipete	Q1	10/03/2026	Done	55	100
049. Surabaya Merr	Q1	11/03/2026	Done	90	100
009. Bintaro	Q1	12/03/2026	Done	81	95
020. Surabaya	Q1	12/03/2026	Done	83	100
023. Semarang	Q1	16/03/2026	Done	92	100
003. Depok	Q1	17/03/2026	Done	61	100
016. Cibubur	Q1	26/03/2026	Done	69	100
033. Taman Mini	Q2	01/04/2026	Done	80	100
032. Cawang	Q2	02/04/2026	Done	78	100
053. Alam Sutera	Q2	08/04/2026	Done	71	100
036. Jagakarsa	Q2	09/04/2026	Done	70	100
024. Grand Wisata	Q2	13/04/2026	Done	57	100
027. Rawamangun	Q2	14/04/2026	Done	58	100
026. Tangerang City	Q2	15/04/2026	Done	75	100
042. Muara Karang	Q2	16/04/2026	Done	82	100
044. Depok GDC	Q2	16/04/2026	Done	67	100
008. PIK	Q2	23/04/2026	Done	71	100
030. Ciputat	Q2	23/04/2026	Done	85	100
028. Taman Palem	Q2	24/04/2026	Done	57	100
055. Gajah Mada	Q2	28/04/2026	Done	83	100
046. Kaliurang Jogja	Q1	30/04/2026	Done	76	100
004. Cipete	Q2	30/04/2026	Done	54	100
038. Greenlake	Q2	30/04/2026	Done	73	100
047. Bogor Tajur	Q2	30/04/2026	Done	83	100
035. Ciledug	Q2	06/05/2026	Done	58	94
048. PIK 2	Q2	06/05/2026	Done	81	100
011. Kuningan	Q2	07/05/2026	Done	70	100
005. Greenville	Q2	08/05/2026	Done	66	100
025. Bogor Baru	Q2	12/05/2026	Done	51	100
031. Sawangan	Q2	13/05/2026	Done	76	100
003. Depok	Q2	20/05/2026	Done	50	100
059. Malang Suhat	Q2	25/05/2026	Done	68	100
056. Citra Garden 2	Q2	26/05/2026	Done	63	100
057. Cikupa	Q2	28/05/2026	Done	71	100
021. Kelapa Gading Boulevard	Q2	02/06/2026	Done	69	100
043. Tebet	Q2	03/06/2026	Done	69	70
007. Bekasi	Q2	04/06/2026	Done	50	100
060. Palmerah	Q2	05/06/2026	Done	63	100
010. Bogor	Q2	09/06/2026	Done	72	100
016. Cibubur	Q2	10/06/2026	Done	66	100
054. Bogor Gunung Batu	Q2	11/06/2026	Done	78	100
051. Bintaro Sektor 9	Q2	17/06/2026	Done	64	100
018. Harapan Indah	Q2	29/06/2026	Done	43	82
050. Cibinong	Q3	06/07/2026	Done	71	100
034. Puri Indah	Q2	09/07/2026	Done	63	100
034. Puri Indah	Q3	09/07/2026	Done	63	100
045. FX Mall	Q3	20/07/2026	Done	79	100
052. BSD Rawa Buntu	Q3	21/07/2026	Done	81	100
019. Cinere	Q3	22/07/2026	Done	65	100
031. Sawangan	Q3	24/07/2026	Done	82	100`;

let sql = '';
data.split('\n').forEach(line => {
  const parts = line.split('\t');
  if(parts.length < 5) return;
  const [branch, period, dateStr, status, fc, spv] = parts;
  const branch_name = branch.trim();
  const p = period.trim();
  const [d,m,y] = dateStr.trim().split('/');
  const date = y+'-'+m.padStart(2,'0')+'-'+d.padStart(2,'0');
  const f = fc && fc.trim() ? fc.trim() : 'NULL';
  const s = spv && spv.trim() ? spv.trim() : 'NULL';
  const bn = branch_name.replace(/\\/g, '\\\\').replace(/'/g, "''");
  sql += `UPDATE inspection_reports SET fc_score = ${f}, spv_score = ${s} WHERE branch_id = (SELECT id FROM branches WHERE full_name = '${bn}' LIMIT 1) AND inspection_date = '${date}';\n`;
});
fs.writeFileSync('update_scores_pure3.sql', sql, 'utf8');
