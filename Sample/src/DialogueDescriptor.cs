using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

public class DialogueTable
{
	[JsonProperty("Dialogue")]
	public List<DialogueDescriptor> Dialogues { get; set; }
}

public class DialogueDescriptor
{
	[JsonProperty("ID")]
	public string ID { get; set; }

	[JsonProperty("Description")]
	public string Description { get; set; }
}
