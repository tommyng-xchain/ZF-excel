export let config: any = {
  // Table__model= jq("#Table_id__model").DataTable({
  scrollCollapse: true,
  paging: false,

  searchPanes: true,
  select: {
    style: "multi",
  },
  deferRender: true,
  data: null,
  columns: [
    // benz_prototypemodeldesignationid
    { title: "Model Name", name: "benz_name" },
    { title: "Type Class", name: "_benz_prototypemodeldesignationtypeclass_value" },
    { title: "ICE / BEV", name: "benz_icebevhybrid" },
    { title: "AMG / Non-AMG", name: "benz_amgmaybachna" },
  ],
  columnDefs: [
    {
      targets: [1],
      searchPanes: {
        show: true,
        orthogonal: "searchpanes",
        collapse: true,
        initCollapsed: true,
      },
    },
    {
      targets: [2],
      searchPanes: {
        show: true,
        orthogonal: "searchpanes",
        collapse: true,
        initCollapsed: true,
      },
    },
    {
      targets: [3],
      searchPanes: {
        show: true,
        orthogonal: "searchpanes",
        collapse: true,
        initCollapsed: true,
      },
    },
  ],

  dom: "Pfrtip",
};
