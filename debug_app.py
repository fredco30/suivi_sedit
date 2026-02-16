"""
Debug pour comprendre pourquoi le cache ne fonctionne pas dans l'application
"""

import os
import time
from marches_module import MarchesAnalyzer

def debug_application():
    print("=" * 80)
    print("DEBUG: SIMULATION DU COMPORTEMENT DE L'APPLICATION")
    print("=" * 80)

    # Vérifier les fichiers
    cache_path = 'marches_cache.db'
    db_path = 'suivi_commandes.db'
    excel_path = 'liste fact.xls'

    print("\n📁 Vérification des fichiers:")
    print(f"   Cache SQLite: {'✅' if os.path.exists(cache_path) else '❌'} {cache_path}")
    if os.path.exists(cache_path):
        stat = os.stat(cache_path)
        print(f"      Taille: {stat.st_size} octets")
        print(f"      Modifié: {time.ctime(stat.st_mtime)}")

    print(f"   Base de données: {'✅' if os.path.exists(db_path) else '❌'} {db_path}")
    print(f"   Fichier Excel: {'✅' if os.path.exists(excel_path) else '❌'} {excel_path}")

    if os.path.exists(excel_path):
        stat = os.stat(excel_path)
        print(f"      Taille: {stat.st_size} octets")
        print(f"      Modifié: {time.ctime(stat.st_mtime)}")

    # Simuler ce que fait l'application
    print("\n" + "=" * 80)
    print("SIMULATION 1: Comme dans refresh_marches_data()")
    print("=" * 80)

    fact_path = excel_path
    print(f"📋 Utilisation du fichier Excel: {fact_path}")

    print("\n🔄 Création de MarchesAnalyzer (comme dans l'app)...")
    print(f"   excel_path = {fact_path}")
    print(f"   database = None")
    print(f"   use_cache = (défaut = True)")

    start = time.time()
    analyzer = MarchesAnalyzer(fact_path, database=None)
    print(f"   ✅ Analyzer créé")
    print(f"   use_cache = {analyzer.use_cache}")
    print(f"   sync = {analyzer.sync}")

    print("\n🔄 Appel de load_data()...")
    success = analyzer.load_data()
    duration = time.time() - start

    print(f"\n⏱️  DURÉE: {duration:.3f}s")
    print(f"📊 Succès: {success}")
    if success:
        print(f"   Lignes chargées: {len(analyzer.df_marches)}")

    if analyzer.sync_stats:
        print(f"\n📈 Statistiques de sync:")
        for key, value in analyzer.sync_stats.items():
            print(f"   {key}: {value}")

    # Interpréter les résultats
    print("\n" + "=" * 80)
    print("🔍 ANALYSE")
    print("=" * 80)

    if analyzer.sync_stats:
        status = analyzer.sync_stats.get('status', 'unknown')

        if status == 'cached':
            print("\n✅ LE CACHE A ÉTÉ UTILISÉ!")
            print(f"   Durée: {duration:.3f}s (devrait être < 0.02s)")
            if duration > 0.05:
                print(f"   ⚠️  Mais c'est plus lent que prévu...")
        elif status == 'success':
            nb_inserted = analyzer.sync_stats.get('nb_inserted', 0)
            nb_deleted = analyzer.sync_stats.get('nb_deleted', 0)
            nb_unchanged = analyzer.sync_stats.get('nb_unchanged', 0)

            print(f"\n⚠️  UNE SYNCHRONISATION A ÉTÉ EFFECTUÉE")
            print(f"   Insérées: {nb_inserted}")
            print(f"   Supprimées: {nb_deleted}")
            print(f"   Inchangées: {nb_unchanged}")
            print(f"   Durée: {duration:.3f}s")

            if nb_unchanged > 500 and nb_inserted == 0 and nb_deleted == 0:
                print(f"\n❌ PROBLÈME: Le fichier n'a pas changé mais une sync a été faite!")
                print(f"   Le système de détection ne fonctionne pas correctement")
            else:
                print(f"\n✅ Le fichier a changé, la sync était justifiée")

    # Deuxième lancement pour voir
    print("\n" + "=" * 80)
    print("SIMULATION 2: Deuxième lancement immédiat")
    print("=" * 80)

    time.sleep(1)

    print("\n🔄 Création d'un nouveau MarchesAnalyzer...")
    start = time.time()
    analyzer2 = MarchesAnalyzer(fact_path, database=None)
    success2 = analyzer2.load_data()
    duration2 = time.time() - start

    print(f"\n⏱️  DURÉE: {duration2:.3f}s")
    print(f"📊 Succès: {success2}")

    if analyzer2.sync_stats:
        print(f"\n📈 Statistiques:")
        for key, value in analyzer2.sync_stats.items():
            print(f"   {key}: {value}")

    # Comparaison
    print("\n" + "=" * 80)
    print("📊 COMPARAISON DES DEUX LANCEMENTS")
    print("=" * 80)
    print(f"\n1er lancement: {duration:.3f}s - statut: {analyzer.sync_stats.get('status', 'N/A')}")
    print(f"2ème lancement: {duration2:.3f}s - statut: {analyzer2.sync_stats.get('status', 'N/A')}")

    if duration2 > 0.1:
        print(f"\n❌ PROBLÈME CONFIRMÉ: Le 2ème lancement est trop lent ({duration2:.3f}s)")
        print(f"   Il devrait être < 0.02s si le cache fonctionne")
    else:
        print(f"\n✅ Le cache fonctionne correctement!")

if __name__ == "__main__":
    debug_application()
