using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class AudioManager : MonoBehaviour
{

    public List<AudioClip> bgmList;
    public List<AudioClip> seList;
    // Start is called before the first frame update
    void Start()
    {
        LoadBGM();
        LoadSE();
    }

    // Update is called once per frame
    void Update()
    {
        
    }

    void LoadBGM()
    {
        TextAsset csvFile;
        csvFile = Resources.Load("Audio/BGM") as TextAsset;
        StringReader reader = new StringReader(csvFile.text);

        while (reader.Peek() > -1)
        {
            string line = reader.ReadLine();
            if (line == "")
            {
                break;
            }
            string sourceDir = "Audio";
            string sourcePath = sourceDir + "/" + line;
            bgmList.Add(Resources.Load<AudioClip>(sourcePath));
        }
    }

    void LoadSE()
    {
        TextAsset csvFile;
        csvFile = Resources.Load("Audio/SE") as TextAsset;
        StringReader reader = new StringReader(csvFile.text);

        while (reader.Peek() > -1)
        {
            string line = reader.ReadLine();
            if (line == "")
            {
                break;
            }
            string sourceDir = "Audio";
            string sourcePath = sourceDir + "/" + line;
            seList.Add(Resources.Load<AudioClip>(sourcePath));
        }
    }
}
