using System.Collections.Generic;
using UnityEngine;

public class Example1 : MonoBehaviour
{

    [SerializeField]
    CharacterJobs[] jobs;

    [SerializeField]
    Weapons[] weapons;

    [SerializeField]
    Equipments[] equipments;


#if UNITY_EDITOR

    [SerializeField, HideInInspector]
    string[] filePaths;

    [SerializeField, HideInInspector]
    bool autoRefresh;

#endif

}
