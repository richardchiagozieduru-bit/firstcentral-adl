"""
Database router to direct Subscriber and SubscriberToken model queries to MSSQL database
while keeping all other models (User, Session, UploadSession, Q tasks) on SQLite.

This prevents MSSQL idle connection timeouts during long-running Q Cluster tasks.
"""

class SubscriberDatabaseRouter:
    """
    A router to control database operations for Subscriber-related models.
    
    Routes:
    - Subscriber, SubscriberToken → 'subscribers_db' (MSSQL)
    - All other models → 'default' (SQLite)
    """
    
    # Models that should use the MSSQL subscribers database
    # Note: SubscriberToken stays in default (SQLite) because it has ForeignKey to User
    subscriber_models = {'subscriber'}
    
    def db_for_read(self, model, **hints):
        """
        Direct read operations for Subscriber models to MSSQL
        """
        if model._meta.model_name in self.subscriber_models:
            return 'subscribers_db'
        return 'default'
    
    def db_for_write(self, model, **hints):
        """
        Direct write operations for Subscriber models to MSSQL
        Note: Subscriber is managed=False, so writes won't actually happen
        """
        if model._meta.model_name in self.subscriber_models:
            return 'subscribers_db'
        return 'default'
    
    def allow_relation(self, obj1, obj2, **hints):
        """
        Allow relations between models in the same database
        """
        db_set = {'default', 'subscribers_db'}
        
        # Get the database each object would use
        db1 = self.db_for_read(obj1.__class__)
        db2 = self.db_for_read(obj2.__class__)
        
        # Allow if both in known databases
        if db1 in db_set and db2 in db_set:
            return True
        return None
    
    def allow_migrate(self, db, app_label, model_name=None, **hints):
        """
        Ensure migrations are applied to the correct database
        
        - Subscriber models: Only migrate on 'subscribers_db' (though managed=False prevents this)
        - All other models: Only migrate on 'default'
        """
        if model_name in self.subscriber_models:
            # Subscriber models should only exist in subscribers_db
            # But since managed=False, Django won't create tables anyway
            return db == 'subscribers_db'
        
        # All other models should only be on default database
        return db == 'default'
