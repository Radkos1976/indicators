
async function container_wyd (Labels,Title,Subtitle,DataSets) {
  const json_cont = {
      error : "no",
      labels : Labels,
      title : Title,
      subtitle : Subtitle,
      datasets : [
        {
          dt_nam:"Współczynnik wydajności",
          fillColor : "rgba(255, 79, 48, 0.3)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "%",
          j_nam : "Procent",
          data : DataSets[0]
        }],
      next_datasets:[
        {
          dt_nam:"Wykonanie godz.",
          fillColor : "rgba(215, 44, 44, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "h",
          j_nam : "Godziny",
          data : DataSets[1]
        },
        {
          dt_nam:"Dostępne godz. brutto",
          fillColor : "rgba(61, 98, 66, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "h",
          j_nam : "Godziny brutto",
          data : DataSets[2]
        },
        {
          dt_nam:"Dostępne godz. netto",
          fillColor : "rgba(71, 98, 66, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "h",
          j_nam : "Godziny netto",
          data : DataSets[3]
        }
      ],
      menu_item :[
        { capt: 'Współczynnik wydajności',
          dtaset: 0,
        },
        { capt: 'Wielkość wykonania godz.',
          dtaset: 1,
        },
        { capt: 'Dostępne godz. razem',
          dtaset: 2,
        },
        { capt: 'Dostępne godz. bez przerw',
          dtaset: 3,
        }
      ],    
    };
  return json_cont;
}

module.exports = {
  Create: (Labels,Title,Subtitle,DataSets) => container_wyd(Labels,Title,Subtitle,DataSets)
};
