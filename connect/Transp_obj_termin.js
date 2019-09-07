async function container_term (Labels,Title,Subtitle,DataSets) {
  const json_cont = {
      error : "no",
      labels : Labels,
      title : Title,
      subtitle : Subtitle,
      datasets : [
        {
          dt_nam:"Współczynnik terminowości",
          fillColor : "rgba(180,152,197,0.5)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "%",
          j_nam : "Procent",
          data : DataSets[0]
        }],
        next_datasets : [{
          dt_nam:"Ilość zaplanowana",
          fillColor : "rgba(215, 44, 44, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "szt",
          j_nam : "Sztuki",
          data : DataSets[1]
        },
        {
          dt_nam:"Wykonanie na czas",
          fillColor : "rgba(61, 98, 66, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "szt",
          j_nam : "Sztuki",
          data : DataSets[2]
        },
        {
          dt_nam:"Wykonanie po czasie",
          fillColor : "rgba(71, 98, 66, 0.6)",
          strokeColor : "rgba(151,187,205,1)",
          pointColor : "rgba(220,220,220,1)",
          pointStrokeColor : "#fff",
          jedn : "szt",
          j_nam : "Sztuki",
          data : DataSets[3]
        }
      ],
      menu_item :[
        { capt: 'Współczynnik terminowości',
          dtaset: 0,
        },
        { capt: 'Ilość zaplanowana',
          dtaset: 1,
        },
        { capt: 'Wykonanie na czas',
          dtaset: 2,
        },
        { capt: 'Wykonanie za późno',
          dtaset: 3,
        }
      ],    
    };
  return json_cont;
}

module.exports = {
  Create: (Labels,Title,Subtitle,DataSets) => container_term(Labels,Title,Subtitle,DataSets)
};
