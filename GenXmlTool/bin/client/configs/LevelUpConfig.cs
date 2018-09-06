// Auto Generated Code
// Author roy

using System.Collections.Generic;
using System.Xml;

public class LevelUpConfig
{
	public int Level;
	public int PlayerLevelUpExp;
	public string CardLevelUpRes;
	public string CardDecomposeRes;

	public static readonly string urlKey = "LevelUpConfig";
	static Dictionary<int,LevelUpConfig> AllDatas;

	public static void Parse(XmlNode node)
	{
		AllDatas = new Dictionary<int,LevelUpConfig>();
		if (node != null)
		{
			XmlNodeList nodeList = node.ChildNodes;
			if (nodeList != null && nodeList.Count > 0)
			{
				foreach (XmlElement el in nodeList)
				{
					LevelUpConfig config = new LevelUpConfig();

					int.TryParse(el.GetAttribute ("Level"), out config.Level);

					int.TryParse(el.GetAttribute ("PlayerLevelUpExp"), out config.PlayerLevelUpExp);

					config.CardLevelUpRes = el.GetAttribute ("CardLevelUpRes");

					config.CardDecomposeRes = el.GetAttribute ("CardDecomposeRes");

					AllDatas.Add(config.Level, config);
				}
			}
		}
	}

	public static LevelUpConfig Get(int key)
	{
		if (AllDatas != null && AllDatas.ContainsKey(key))
			return AllDatas[key];
		return null;
	}

	public static Dictionary<int,LevelUpConfig> Get()
	{
		return AllDatas;
	}
}
