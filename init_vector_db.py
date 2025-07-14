from vector_rag_db import VectorRAGDatabase
import logging

DOCUMENTS_DIR = r"C:\Users\Гульшат\Desktop\ДИПЛОМ25\документы"
VECTOR_DB_PATH = r"C:\Users\Гульшат\Desktop\ДИПЛОМ25\векторная_база"

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    
    # Инициализация базы
    vector_db = VectorRAGDatabase(DOCUMENTS_DIR, VECTOR_DB_PATH)
    
    # Первичное создание или полное обновление
    vector_db.index_documents()  
    
    # Для инкрементного обновления используйте:
    # vector_db.update_documents()
