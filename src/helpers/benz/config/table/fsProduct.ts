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
      // benz_fsproductid
      { title: "FS Product Name", name: "benz_name" },
      { title: "type Name", name: "benz_nameoffinanceproduct" },
      { title: "type ID", name: "_benz_financeproduct_value" }
    ],
    columnDefs: [],
  
    dom: "Pfrtip",
  };
  