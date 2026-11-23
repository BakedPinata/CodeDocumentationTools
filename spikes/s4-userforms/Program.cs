using S4Spike;

return args switch
{
  ["inspect", var path] => Inspector.Run(path),
  ["build", var sample, var output, .. var only] => Builder.Run(sample, output, only),
  ["dir", var workbook] => DirDump.Run(workbook),
  ["compare", var workbook, var frx] => Compare.Run(workbook, frx),
  _ => Usage(),
};

static int Usage()
{
  Console.Error.WriteLine("s4spike build <sample folder> <output folder> [F1 F2 ...]");
  Console.Error.WriteLine("s4spike inspect <workbook.xlsm | form.frx>");
  Console.Error.WriteLine("s4spike compare <workbook.xlsm> <form.frx>");
  return 2;
}
